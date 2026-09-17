import JSZip from 'jszip';
import { getSchema, type JSONContent } from '@tiptap/core';
import type { OfficeFile } from './model';
import { wordJSON, wordExtensions } from './word-extensions';
import { validateSectionState } from './word-section-breaks';
import { prepareLiveSections } from './docx-live-sections';
import { collapsedWordNavigationAt } from './docx-navigation-bookmark';
import { inspectZip, sanitizeHTML } from './formats';
import { defaultFont } from './fonts';
import { paragraphBreaks } from './paragraph-layout';
import { wordLineSpacing } from './word-line-spacing';
import { wordParagraphSpaceFromAttrs, wordParagraphSpaceXml } from './word-paragraph-spacing';
import { wordTabStops, wordDecimalSymbol } from './word-tab-stops';
import { docxLinks, type DocxLinks } from './docx-links';
import { docxStructureSchema, liveSectionProperties } from './docx-sections';
import {
  coalescePlainWordRuns,
  isPlainWordRun,
  normalizeWordTextSpace,
} from './docx-run-normalization';
import { wordEditSession } from './docx-edit-session';
import { initializeWordSourceSession } from './docx-initial-session';
import { registerMissingWordSessions } from './docx-missing-session-register';
import { wordRunSession } from './word-edit-run';
import { writeWordStoryOptions } from './docx-story-options';
import { prepareWordStoryClones, writeWordStoryReferences } from './docx-story-references';
import { readDocx, WORD_NS, child, descendants, wElement, val } from './docx-import';
import { writeWordFontFeatures, wordFontFeaturesXml, WORD_2010_NS } from './word-font-features';
import { wordScriptValue } from './word-script';
import { writeWordText } from './word-xml';
import { materializeWordFinalSection } from './word-section-defaults';
import { wordHyphenText } from './word-hyphen';

const equal = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
const paragraph = (node: JSONContent) => ['paragraph', 'heading'].includes(node.type || '');
const keyOf = (node: JSONContent) => node.attrs?.sourceParagraph as string | null;
const deletionElements = new Set(
  'pPr r rPr t br tab softHyphen tabs pStyle jc spacing ind bidi b i u sz szCs rFonts color lang kern keepNext keepLines pageBreakBefore widowControl lastRenderedPageBreak'.split(
    ' ',
  ),
);
function deletableFontFeature(element: Element) {
  if (
    element.namespaceURI !== WORD_2010_NS ||
    !['ligatures', 'cntxtAlts'].includes(element.localName) ||
    element.childNodes.length ||
    element.parentElement?.namespaceURI !== WORD_NS ||
    element.parentElement.localName !== 'rPr' ||
    [...element.attributes].some(
      (attr) =>
        attr.namespaceURI !== 'http://www.w3.org/2000/xmlns/' &&
        !(attr.namespaceURI === WORD_2010_NS && attr.localName === 'val'),
    )
  )
    return false;
  try {
    wordFontFeaturesXml(element.parentElement);
    return true;
  } catch {
    return false;
  }
}
function deletableScript(element: Element) {
  return (
    element.namespaceURI === WORD_NS &&
    element.localName === 'vertAlign' &&
    element.parentElement?.namespaceURI === WORD_NS &&
    element.parentElement.localName === 'rPr' &&
    !element.childNodes.length &&
    wordScriptValue(val(element)) !== null &&
    [...element.attributes].every(
      (attr) =>
        attr.namespaceURI === 'http://www.w3.org/2000/xmlns/' ||
        (attr.namespaceURI === WORD_NS && attr.localName === 'val'),
    )
  );
}
function paragraphDeletionStructures(p: Element) {
  const navigation = new Set<Element>();
  for (const start of [...p.children]) {
    if (
      start.namespaceURI !== WORD_NS ||
      start.localName !== 'bookmarkStart' ||
      val(start, 'name') !== '_GoBack' ||
      start.children.length
    )
      continue;
    const end = start.nextElementSibling,
      id = val(start, 'id');
    if (
      end?.namespaceURI === WORD_NS &&
      end.localName === 'bookmarkEnd' &&
      !end.children.length &&
      /^\d+$/.test(id) &&
      val(end, 'id') === id &&
      descendants(p.ownerDocument, 'bookmarkStart').filter((e) => val(e, 'id') === id).length ===
        1 &&
      descendants(p.ownerDocument, 'bookmarkEnd').filter((e) => val(e, 'id') === id).length === 1
    ) {
      navigation.add(start);
      navigation.add(end);
    }
  }
  return [...p.getElementsByTagName('*')].filter(
    (e) =>
      !navigation.has(e) &&
      !deletableFontFeature(e) &&
      !deletableScript(e) &&
      (e.namespaceURI !== WORD_NS ||
        !deletionElements.has(e.localName) ||
        (e.localName === 'softHyphen' &&
          (e.childNodes.length > 0 ||
            e.parentElement?.localName !== 'r' ||
            [...e.attributes].some(
              (attr) => attr.namespaceURI !== 'http://www.w3.org/2000/xmlns/',
            ))) ||
        (e.parentElement === p && !['pPr', 'r'].includes(e.localName))),
  );
}
function fail(detail: string): never {
  throw new Error(
    `This DOCX edit cannot yet be exported without losing document features: ${detail}. Save a Noffice backup to keep your edits.`,
  );
}
const serial = (node: Node) => new XMLSerializer().serializeToString(node);
type Marks = NonNullable<JSONContent['marks']>;
function flat(node: JSONContent): { text: string; marks: Marks[] } {
  let text = '';
  const marks: Marks[] = [];
  for (const run of node.content || []) {
    if (
      !['text', 'hardBreak', 'wordTab', 'wordHyphen', 'wordPageBreak', 'wordColumnBreak'].includes(
        run.type || '',
      )
    )
      fail('editing this paragraph’s inline objects is not supported');
    if (run.type === 'text' && run.text?.includes('\u001f'))
      fail('an unsupported control character');
    const value =
      run.type === 'wordHyphen'
        ? (wordHyphenText(run.attrs?.kind, 'native') ?? fail('an invalid hyphen character'))
        : run.type === 'hardBreak'
          ? '\n'
          : run.type === 'wordTab'
            ? '\t'
            : run.type === 'wordPageBreak'
              ? '\f'
              : run.type === 'wordColumnBreak'
                ? '\u000e'
                : run.text || '';
    text += value;
    for (let i = 0; i < value.length; i++) marks.push(run.marks || []);
  }
  return { text, marks };
}
function property(parent: Element, name: string) {
  const existing = child(parent, name);
  if (existing) return existing;
  const el = wElement(parent.ownerDocument, name);
  const orders: Record<string, string[]> = {
    pPr: 'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange'.split(
      ' ',
    ),
    rPr: 'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange'.split(
      ' ',
    ),
    sectPr:
      'headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange'.split(
        ' ',
      ),
  };
  const order = orders[parent.localName] || [];
  const at = order.indexOf(name);
  parent.insertBefore(
    el,
    [...parent.children].find((e) => at >= 0 && order.indexOf(e.localName) > at) || null,
  );
  return el;
}
function props(p: Element) {
  let pr = child(p, 'pPr');
  if (!pr) {
    pr = wElement(p.ownerDocument, 'pPr');
    p.prepend(pr);
  }
  return pr;
}
type ParagraphMarkSources = Map<string, { owner: Document; properties: Element | null }>;
function applyParagraphMarkSource(p: Element, node: JSONContent, sources: ParagraphMarkSources) {
  const key = node.attrs?.paragraphMarkSource;
  if (key == null) return;
  const source = key === 'new' ? null : sources.get(key);
  if (key !== 'new' && (!source || source.owner !== p.ownerDocument))
    fail('an invalid or cross-part paragraph mark source');
  const pr = props(p);
  child(pr, 'rPr')?.remove();
  if (source?.properties) property(pr, 'rPr').replaceWith(source.properties.cloneNode(true));
}
const setVal = (el: Element, value: string, attr = 'val') =>
  el.setAttributeNS(WORD_NS, `w:${attr}`, value);
const twips = (value: string) =>
  String(Math.round(parseFloat(value) * (value.endsWith('pt') ? 20 : 15)));
function paragraphProperties(p: Element, before: JSONContent, after: JSONContent) {
  const previous = before.attrs || {},
    next = after.attrs || {};
  for (const name of new Set([...Object.keys(previous), ...Object.keys(next)])) {
    if (
      equal(previous[name], next[name]) ||
      name === 'sourceParagraph' ||
      name === 'paragraphMarkSource' ||
      name === 'level'
    )
      continue;
    // These reading hints describe document settings and validation, not
    // paragraph XML. Document-wide values are validated before patching.
    if (
      [
        'paragraphDefaultTab',
        'paragraphDecimalSymbol',
        'paragraphTabUnsupported',
        'paragraphStyleId',
      ].includes(name)
    )
      continue;
    if (name === 'paragraphTabs') {
      if (previous.paragraphTabUnsupported || next.paragraphTabUnsupported)
        fail('changing unsupported tab stops');
      const oldStops = wordTabStops(previous[name] ?? []),
        newStops = wordTabStops(next[name] ?? []);
      if (!oldStops || !newStops) fail('invalid tab stops');
      const wanted = new Set(newStops.map((stop) => stop.position));
      const tabs = property(props(p), 'tabs');
      // Earlier clears may hide inherited stops that are absent from the
      // effective list. Retain those clears or editing another stop revives them.
      const cleared = new Set(
        oldStops.filter((stop) => !wanted.has(stop.position)).map((stop) => stop.position),
      );
      for (const tab of tabs.children) {
        if (tab.namespaceURI !== WORD_NS || tab.localName !== 'tab') fail('unknown tab definition');
        if (val(tab) !== 'clear') continue;
        const position = Number(val(tab, 'pos'));
        if (!/^[+-]?\d+$/.test(val(tab, 'pos')) || !Number.isSafeInteger(position))
          fail('invalid retained tab clear');
        if (!wanted.has(position)) cleared.add(position);
      }
      tabs.replaceChildren(
        ...[...cleared]
          .sort((a, b) => a - b)
          .map((position) =>
            wElement(p.ownerDocument, 'tab', { val: 'clear', pos: String(position) }),
          ),
        ...newStops.map((stop) =>
          wElement(p.ownerDocument, 'tab', {
            val: stop.alignment,
            pos: String(stop.position),
            ...(stop.leader === 'none' ? {} : { leader: stop.leader }),
          }),
        ),
      );
    } else if (name === 'paragraphFontFeatures') {
      writeWordFontFeatures(property(props(p), 'rPr'), next[name] ?? 0);
    } else if (name === 'paragraphScript') {
      const value = next[name] == null ? 'baseline' : wordScriptValue(next[name]);
      if (!value) fail('invalid paragraph script');
      setVal(property(property(props(p), 'rPr'), 'vertAlign'), value!);
    } else if (name === 'textAlign')
      setVal(property(props(p), 'jc'), next[name] === 'justify' ? 'both' : next[name] || 'left');
    else if (name === 'direction')
      setVal(property(props(p), 'bidi'), next[name] === 'rtl' ? '1' : '0');
    else if (
      name === 'paragraphFontSize' ||
      name === 'paragraphFontFamily' ||
      name === 'paragraphKerning'
    ) {
      const mark = property(props(p), 'rPr');
      if (name === 'paragraphKerning') {
        const threshold = next[name] ?? 0;
        if (!Number.isInteger(threshold) || Number(threshold) < 0 || Number(threshold) > 3276)
          fail('invalid paragraph kerning threshold');
        setVal(property(mark, 'kern'), String(threshold));
      } else if (name === 'paragraphFontSize') {
        const size = parseFloat(next[name] || '10') * (String(next[name]).endsWith('px') ? 1.5 : 2);
        if (!Number.isFinite(size) || size < 1 || size > 3276)
          fail('unsupported paragraph font size');
        setVal(property(mark, 'sz'), String(Math.round(size)));
      } else {
        const fonts = property(mark, 'rFonts');
        for (const key of ['ascii', 'hAnsi']) {
          setVal(fonts, String(next[name] || ''), key);
          fonts.removeAttributeNS(WORD_NS, key + 'Theme');
        }
      }
    } else if (name === 'paragraphContextualSpacing')
      setVal(property(props(p), 'contextualSpacing'), next[name] ? '1' : '0');
    else if (paragraphBreaks.some(([key]) => key === name))
      setVal(property(props(p), name), next[name] ? '1' : '0');
    else if (['indentStart', 'indentEnd', 'firstLineIndent'].includes(name)) {
      const amount = next[name] ? Number(twips(next[name])) : 0;
      if (!Number.isFinite(amount) || Math.abs(amount) > 31680)
        fail('paragraph indentation exceeds 1584 points');
      const indent = property(props(p), 'ind');
      if (name === 'firstLineIndent') {
        if (amount >= 0 && parseFloat(previous[name] || '0') < 0 && !val(indent, 'hanging'))
          fail(
            'changing an inherited hanging indent to a first-line indent requires style editing',
          );
        for (const attr of ['firstLineChars', 'hangingChars']) setVal(indent, '0', attr);
        setVal(indent, String(Math.max(0, amount)), 'firstLine');
        setVal(indent, String(Math.max(0, -amount)), 'hanging');
        // A nonzero hanging indent overrides firstLine; zero removes that override.
        if (amount >= 0) indent.removeAttributeNS(WORD_NS, 'hanging');
      } else {
        const keys = name === 'indentStart' ? ['start', 'left'] : ['end', 'right'];
        for (const attr of keys) {
          setVal(indent, '0', attr + 'Chars');
          setVal(indent, String(amount), attr);
        }
      }
    } else if (
      [
        'paragraphLineHeight',
        'paragraphLineRule',
        'spaceBefore',
        'spaceAfter',
        'paragraphSpaceBefore',
        'paragraphSpaceAfter',
      ].includes(name)
    ) {
      const spacing = property(props(p), 'spacing');
      if (name === 'paragraphLineHeight' || name === 'paragraphLineRule') {
        const spec = wordLineSpacing(next.paragraphLineHeight || '1', next.paragraphLineRule);
        if (!spec) fail('unsupported line spacing');
        setVal(spacing, String(spec.line), 'line');
        setVal(spacing, spec.rule, 'lineRule');
      } else {
        const side = name.endsWith('Before') ? 'before' : 'after';
        const space = wordParagraphSpaceFromAttrs(next, side);
        if (!space) fail('invalid paragraph spacing');
        for (const [key, value] of Object.entries(wordParagraphSpaceXml(side, space)))
          setVal(spacing, value, key);
      }
    } else fail(`unsupported paragraph property ${name}`);
  }
  if (before.type !== after.type || previous.level !== next.level) {
    // Style IDs are package-defined; never invent a reference to a missing heading style.
    fail('changing imported heading styles is not supported');
  }
}
function markStyle(marks: Marks): Record<string, unknown> {
  const style: Record<string, unknown> = {};
  for (const mark of marks) {
    if (mark.type === 'wordEditRun') continue;
    if (mark.type === 'textStyle') Object.assign(style, mark.attrs);
    else style[mark.type] = mark.attrs || true;
  }
  return Object.fromEntries(
    Object.entries(style).filter(
      ([, value]) => value !== null && value !== undefined && value !== '',
    ),
  );
}
function color(value: unknown) {
  const text = String(value || '').trim();
  const rgb = /^rgb\(\s*(\d+)[, ]+\s*(\d+)[, ]+\s*(\d+)\s*\)$/.exec(text);
  if (rgb)
    return rgb
      .slice(1)
      .map((n) => Math.min(255, Number(n)).toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase();
  if (/^#[a-f\d]{6}$/i.test(text)) return text.slice(1).toUpperCase();
  if (/^#[a-f\d]{3}$/i.test(text))
    return [...text.slice(1)]
      .map((c) => c + c)
      .join('')
      .toUpperCase();
  return text === 'yellow'
    ? 'FFFF00'
    : text === 'black'
      ? '000000'
      : text === 'white'
        ? 'FFFFFF'
        : 'auto';
}
function applyMarks(
  run: Element,
  before: Marks,
  after: Marks,
  paragraphAttrs: JSONContent['attrs'] = {},
) {
  const old = markStyle(before),
    next = markStyle(after);
  const changed = [...new Set([...Object.keys(old), ...Object.keys(next)])].filter(
    (key) => !equal(old[key], next[key]),
  );
  if (!changed.length) return;
  let pr = child(run, 'rPr');
  if (!pr) {
    pr = wElement(run.ownerDocument, 'rPr');
    run.prepend(pr);
  }
  for (const name of changed) {
    if (name === 'link') continue; // A link wraps runs; its destination is a part relationship.
    const on = !!next[name];
    const boolean = ({ bold: 'b', italic: 'i', strike: 'strike' } as Record<string, string>)[name];
    if (boolean) {
      setVal(property(pr, boolean), on ? '1' : '0');
      if (boolean === 'b' || boolean === 'i') setVal(property(pr, `${boolean}Cs`), on ? '1' : '0');
    } else if (name === 'underline') setVal(property(pr, 'u'), on ? 'single' : 'none');
    else if (name === 'subscript' || name === 'superscript')
      setVal(
        property(pr, 'vertAlign'),
        next.superscript ? 'superscript' : next.subscript ? 'subscript' : 'baseline',
      );
    else if (name === 'fontFamily') {
      const fonts = property(pr, 'rFonts');
      for (const attr of ['asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme'])
        fonts.removeAttributeNS(WORD_NS, attr);
      for (const attr of ['ascii', 'hAnsi', 'eastAsia', 'cs'])
        setVal(
          fonts,
          String(next[name] || paragraphAttrs.paragraphFontFamily || defaultFont).replace(
            /^['"]|['"]$/g,
            '',
          ),
          attr,
        );
    } else if (name === 'fontSize') {
      const size = String(next[name] || paragraphAttrs.paragraphFontSize || '12pt');
      const halfPoints = Math.round(parseFloat(size) * (size.endsWith('px') ? 1.5 : 2));
      if (!Number.isFinite(halfPoints) || halfPoints < 1 || halfPoints > 3276)
        fail('invalid font size');
      setVal(property(pr, 'sz'), String(halfPoints));
      setVal(property(pr, 'szCs'), String(halfPoints));
    } else if (name === 'wordFontFeatures') {
      writeWordFontFeatures(pr, next[name] ?? 0);
    } else if (name === 'wordKerning') {
      const threshold = next[name] ?? 0;
      if (!Number.isInteger(threshold) || Number(threshold) < 0 || Number(threshold) > 3276)
        fail('invalid font kerning threshold');
      setVal(property(pr, 'kern'), String(threshold));
    } else if (name === 'color') {
      const c = property(pr, 'color');
      const nextColor = color(next[name]);
      if (
        val(c).toLowerCase() === nextColor.toLowerCase() &&
        !['themeColor', 'themeTint', 'themeShade'].some((attr) => c.hasAttributeNS(WORD_NS, attr))
      )
        continue;
      for (const attr of ['themeColor', 'themeTint', 'themeShade'])
        c.removeAttributeNS(WORD_NS, attr);
      setVal(c, nextColor);
    } else if (name === 'highlight' || name === 'backgroundColor') {
      child(pr, 'highlight')?.remove();
      const shd = property(pr, 'shd');
      setVal(shd, 'clear');
      setVal(
        shd,
        color(name === 'highlight' ? (next[name] as { color?: string })?.color : next[name]),
        'fill',
      );
    } else fail(`unsupported text formatting ${name}`);
  }
}

/** Keep the source run properties, changing only properties the user changed in the editor. */
function patchParagraph(
  p: Element,
  before: JSONContent,
  after: JSONContent,
  links: DocxLinks,
  insertedRuns: Set<Element>,
  editId: (session?: string) => string | undefined,
  markSources: ParagraphMarkSources,
) {
  applyParagraphMarkSource(p, after, markSources);
  paragraphProperties(p, before, after);
  if (equal(before.content, after.content)) return;
  const old = flat(before),
    next = flat(after);
  if (old.text.length > 200000 || next.text.length > 200000)
    fail('this paragraph exceeds the editing limit');
  const textValue = (e: Element) =>
    e.localName === 'br'
      ? val(e, 'type') === 'page'
        ? '\f'
        : val(e, 'type') === 'column'
          ? '\u000e'
          : '\n'
      : e.localName === 'softHyphen'
        ? '\u001f'
        : e.localName === 'tab'
          ? '\t'
          : e.textContent || '';
  const texts = Array.from(p.getElementsByTagNameNS(WORD_NS, '*'))
    .filter((e) => ['t', 'br', 'tab', 'softHyphen'].includes(e.localName))
    .filter((t) => {
      let parent = t.parentElement;
      while (parent && parent !== p) {
        // w:tab also names a stop definition in pPr/tabs. Only inline
        // characters contribute to the semantic paragraph text.
        if (['p', 'pPr', 'rPr', 'del', 'moveFrom', 'txbxContent'].includes(parent.localName))
          return false;
        parent = parent.parentElement;
      }
      return parent === p;
    });
  if (texts.map(textValue).join('') !== old.text)
    fail('this paragraph contains fields, references or text the editor cannot map precisely');
  let start = 0;
  while (
    start < Math.min(old.text.length, next.text.length) &&
    old.text[start] === next.text[start]
  )
    start++;
  let tail = 0;
  while (
    tail < old.text.length - start &&
    tail < next.text.length - start &&
    old.text[old.text.length - 1 - tail] === next.text[next.text.length - 1 - tail]
  )
    tail++;
  const end = old.text.length - tail,
    insert = next.text.slice(start, next.text.length - tail);
  const plainPrefix =
    start === 0 && end === 0 && insert.length > 0 && !/[\t\n\f\u000e\u001f]/.test(insert);
  const plainReplacement =
    start === 0 && end > 0 && insert.length > 0 && !/[\t\n\f\u000e\u001f]/.test(insert);
  const fixedPrefix: Element[] = [];
  const movedNavigation: Element[] = [];
  let navigationOffset = insert.length;
  if (plainPrefix) {
    // Word can store a whole-paragraph navigation selection with its end
    // immediately after w:p, covering the paragraph mark. Typing a prefix
    // collapses that transient _GoBack range after the inserted text.
    // Admit only this independently observed shape and unique identity.
    const start = [...p.children].find((e) => e.localName !== 'pPr');
    const end = p.nextElementSibling;
    const id = start && val(start, 'id');
    const attrs = (element: Element, names: string[]) =>
      !element.childNodes.length &&
      [...element.attributes].every(
        (attr) =>
          attr.namespaceURI === 'http://www.w3.org/2000/xmlns/' ||
          (attr.namespaceURI === WORD_NS && names.includes(attr.localName)),
      );
    if (
      start?.namespaceURI === WORD_NS &&
      start.localName === 'bookmarkStart' &&
      val(start, 'name') === '_GoBack' &&
      id &&
      /^\d+$/.test(id) &&
      attrs(start, ['id', 'name']) &&
      end?.namespaceURI === WORD_NS &&
      end.localName === 'bookmarkEnd' &&
      val(end, 'id') === id &&
      attrs(end, ['id']) &&
      p.parentElement?.namespaceURI === WORD_NS &&
      ['body', 'hdr', 'ftr'].includes(p.parentElement.localName) &&
      descendants(p.ownerDocument, 'bookmarkStart').filter((e) => val(e, 'id') === id).length ===
        1 &&
      descendants(p.ownerDocument, 'bookmarkStart').filter((e) => val(e, 'name') === '_GoBack')
        .length === 1 &&
      descendants(p.ownerDocument, 'bookmarkEnd').filter((e) => val(e, 'id') === id).length === 1
    )
      movedNavigation.push(start, end);
  }
  for (const e of [...p.children]) {
    if (e.localName === 'pPr' || fixedPrefix.includes(e)) continue;
    if (movedNavigation.includes(e)) continue;
    // Native keyboard typing keeps a named collapsed anchor before a text
    // prefix, but Word saves _GoBack after that prefix. Consecutive collapsed
    // ranges may be interleaved in the original XML. Validate their complete
    // identities before separating the navigation pair from the named anchors.
    if (
      (start > 0 || plainPrefix || plainReplacement || ['\f', '\u000e'].includes(next.text[0])) &&
      e.namespaceURI === WORD_NS &&
      e.localName === 'bookmarkStart'
    ) {
      const group: Element[] = [];
      for (
        let member: Element | null = e;
        member &&
        member.namespaceURI === WORD_NS &&
        ['bookmarkStart', 'bookmarkEnd'].includes(member.localName);
        member = member.nextElementSibling
      )
        group.push(member);
      const starts = group.filter((member) => member.localName === 'bookmarkStart');
      const counts = new Map<string, number>();
      for (const tag of ['bookmarkStart', 'bookmarkEnd'])
        for (const member of descendants(p.ownerDocument, tag)) {
          const key = tag + ':' + val(member, 'id');
          counts.set(key, (counts.get(key) || 0) + 1);
        }
      const valid =
        group.length === starts.length * 2 &&
        starts.every((member) => {
          const id = val(member, 'id');
          const finish = group.find(
            (candidate) => candidate.localName === 'bookmarkEnd' && val(candidate, 'id') === id,
          );
          return (
            id !== '' &&
            counts.get('bookmarkStart:' + id) === 1 &&
            counts.get('bookmarkEnd:' + id) === 1 &&
            finish &&
            group.indexOf(finish) > group.indexOf(member) &&
            (!plainPrefix ||
              !val(member, 'name').startsWith('_') ||
              val(member, 'name') === '_GoBack') &&
            (!plainReplacement || val(member, 'name') === '_GoBack')
          );
        }) &&
        group.every(
          (member) =>
            !member.children.length &&
            !member.hasAttributeNS(WORD_NS, 'colFirst') &&
            !member.hasAttributeNS(WORD_NS, 'colLast'),
        );
      if (!valid) break;
      const navigation =
        plainPrefix || plainReplacement
          ? starts.find((member) => val(member, 'name') === '_GoBack')
          : undefined;
      for (const member of group)
        (navigation && val(member, 'id') === val(navigation, 'id')
          ? movedNavigation
          : fixedPrefix
        ).push(member);
      continue;
    }
    if (
      e.namespaceURI === WORD_NS &&
      e.localName === 'r' &&
      [...e.children].some((c) => ['footnoteRef', 'endnoteRef'].includes(c.localName)) &&
      [...e.children].every(
        (c) =>
          c.namespaceURI === WORD_NS && ['rPr', 'footnoteRef', 'endnoteRef'].includes(c.localName),
      )
    )
      fixedPrefix.push(e);
    else break;
  }
  const simpleRun = (e: Element) =>
    e.namespaceURI === WORD_NS &&
    e.localName === 'r' &&
    [...e.children].every(
      (c) =>
        c.namespaceURI === WORD_NS &&
        // Last-rendered page markers cache a previous pagination result. The
        // rebuilt edited run invalidates them; explicit w:br remains semantic.
        ['rPr', 't', 'br', 'tab', 'softHyphen', 'lastRenderedPageBreak'].includes(c.localName) &&
        (c.localName !== 'softHyphen' ||
          (!c.childNodes.length &&
            [...c.attributes].every((a) => a.namespaceURI === 'http://www.w3.org/2000/xmlns/'))) &&
        (c.localName !== 'br' ||
          ((!val(c, 'type') || ['textWrapping', 'page', 'column'].includes(val(c, 'type'))) &&
            (!val(c, 'clear') || val(c, 'clear') === 'none'))),
    );
  if (!movedNavigation.length && start === end && start > 0 && insert.length &&
    !/[\t\n\f\u000e\u001f]/.test(insert)) {
    const navigation = collapsedWordNavigationAt(p, start);
    if (navigation.length) {
      movedNavigation.push(...navigation);
      navigationOffset = start + insert.length;
    }
  }
  const simple = [...p.children].every(
    (e) =>
      e.namespaceURI === WORD_NS &&
      (e.localName === 'pPr' ||
        fixedPrefix.includes(e) ||
        movedNavigation.includes(e) ||
        simpleRun(e) ||
        (e.localName === 'hyperlink' && [...e.children].every(simpleRun))),
  );
  const origin = (i: number) =>
    i < start
      ? i
      : i >= start + insert.length
        ? old.text.length - next.text.length + i
        : Math.max(0, Math.min(old.text.length - 1, start < end ? start : start ? start - 1 : 0));
  if (!simple) {
    if (
      texts.some((e) => e.localName !== 't') ||
      next.text.includes('\n') ||
      next.text.includes('\f') ||
      next.text.includes('\u000e') ||
      next.text.includes('\t')
    )
      fail('line-break or tab editing across complex paragraph structures');
    // Text-only changes in bookmark/hyperlink paragraphs retain their exact markup.
    // Review/field boundaries must never be silently moved by a text replacement.
    if (
      [
        'fldChar',
        'fldSimple',
        'ins',
        'del',
        'moveFrom',
        'moveTo',
        'sdt',
        'drawing',
        'pict',
        'footnoteReference',
        'endnoteReference',
      ].some((tag) => descendants(p, tag).length)
    )
      fail('editing text across a field, review structure or embedded object');
    if (next.marks.some((marks, i) => !equal(marks, old.marks[origin(i)] || [])))
      fail('formatting a paragraph with bookmarks or hyperlinks');
    let offset = 0,
      inserted = false;
    for (const text of texts) {
      const value = text.textContent || '',
        right = offset + value.length;
      const leftKeep = value.slice(0, Math.max(0, Math.min(value.length, start - offset)));
      const rightKeep = value.slice(Math.max(0, Math.min(value.length, end - offset)));
      if (right >= start && offset <= end) {
        writeWordText(text, leftKeep + (!inserted ? insert : '') + rightKeep);
        inserted = true;
      }
      offset = right;
    }
    return;
  }
  const sourceRuns: Element[] = [];
  for (const text of texts)
    for (let i = 0; i < textValue(text).length; i++) sourceRuns.push(text.parentElement!);
  const fallback = sourceRuns[0] || wElement(p.ownerDocument, 'r');
  const runIds = new Map<Element, number>();
  for (const run of sourceRuns) if (!runIds.has(run)) runIds.set(run, runIds.size);
  const insertionTemplates = new Set(
    start === end && start > 0 && start < old.text.length && insert.length
      ? [...runIds.keys()].filter(isPlainWordRun)
      : [],
  );
  const newRuns: Element[] = [];
  let previousKey = '',
    activeText: Element | undefined;
  let chunk: string[] = [];
  for (let i = 0; i < next.text.length; i++) {
    if (movedNavigation.length && i === navigationOffset) {
      if (activeText) writeWordText(activeText, chunk.join(''));
      activeText = undefined;
      chunk = [];
      previousKey = '';
      newRuns.push(...movedNavigation);
    }
    const at = origin(i),
      template = sourceRuns[at] || fallback;
    const beforeMarks = old.marks[at] || [],
      afterMarks = next.marks[i] || [];
    if (['\n', '\t', '\f', '\u000e', '\u001f'].includes(next.text[i])) {
      if (activeText) writeWordText(activeText, chunk.join(''));
      activeText = undefined;
      chunk = [];
      previousKey = '';
      const run = template.cloneNode(true) as Element;
      [...run.children].filter((e) => e.localName !== 'rPr').forEach((e) => e.remove());
      applyMarks(run, beforeMarks, afterMarks, after.attrs);
      run.append(
        wElement(
          p.ownerDocument,
          next.text[i] === '\u001f' ? 'softHyphen' : next.text[i] === '\t' ? 'tab' : 'br',
          next.text[i] === '\f'
            ? { type: 'page' }
            : next.text[i] === '\u000e'
              ? { type: 'column' }
              : {},
        ),
      );
      if (next.text[i] === '\u001f') {
        const session = wordRunSession(afterMarks),
          id = session && editId(session);
        if (id) setVal(run, id, 'rsidR');
      }
      links.append(newRuns, run, beforeMarks, afterMarks, template);
      continue;
    }
    const inserted = i >= start && i < start + insert.length && insertionTemplates.has(template);
    const grouping = `${runIds.get(template)}:${JSON.stringify(beforeMarks)}:${JSON.stringify(afterMarks)}:${inserted}`;
    if (grouping !== previousKey || !activeText) {
      if (activeText) writeWordText(activeText, chunk.join(''));
      chunk = [];
      const run = template.cloneNode(true) as Element;
      [...run.children].filter((e) => e.localName !== 'rPr').forEach((e) => e.remove());
      applyMarks(run, beforeMarks, afterMarks, after.attrs);
      const session = wordRunSession(afterMarks);
      if (inserted || session) {
        insertedRuns.add(run);
        const id = editId(session);
        if (id) setVal(run, id, 'rsidR');
      }
      activeText = wElement(p.ownerDocument, 't');
      run.append(activeText);
      links.append(newRuns, run, beforeMarks, afterMarks, template);
      previousKey = grouping;
    }
    chunk.push(next.text[i]);
  }
  if (activeText) writeWordText(activeText, chunk.join(''));
  if (movedNavigation.length && navigationOffset === next.text.length)
    newRuns.push(...movedNavigation);
  [...p.children].filter((e) => e.localName !== 'pPr').forEach((e) => e.remove());
  p.append(...fixedPrefix, ...newRuns);
}

function newParagraph(
  doc: Document,
  node: JSONContent,
  links: DocxLinks,
  editId: (session?: string) => string | undefined,
  markSources: ParagraphMarkSources,
  template?: Element,
  templateNode?: JSONContent,
) {
  const p = wElement(doc, 'p');
  if (template) {
    const pr = child(template, 'pPr');
    if (pr) {
      const copy = pr.cloneNode(true) as Element;
      for (const tag of ['sectPr', 'pPrChange']) child(copy, tag)?.remove();
      p.append(copy);
    }
  }
  if (node.type === 'heading') fail('adding imported heading styles');
  applyParagraphMarkSource(p, node, markSources);
  paragraphProperties(
    p,
    {
      type: 'paragraph',
      attrs: {
        ...Object.fromEntries(Object.keys(node.attrs || {}).map((k) => [k, null])),
        paragraphTabs: templateNode?.attrs?.paragraphTabs ?? null,
        paragraphTabUnsupported: templateNode?.attrs?.paragraphTabUnsupported ?? false,
      },
    },
    node,
  );
  const runs: Element[] = [];
  for (const piece of node.content || []) {
    const run = wElement(doc, 'r');
    applyMarks(run, [], piece.marks || [], node.attrs);
    const session = wordRunSession(piece.marks || []);
    if (session) {
      const id = editId(session);
      if (id) setVal(run, id, 'rsidR');
    }
    if (piece.type === 'text') {
      const text = wElement(doc, 't');
      writeWordText(text, piece.text || '');
      run.append(text);
    } else if (piece.type === 'wordHyphen') {
      const value = wordHyphenText(piece.attrs?.kind, 'native');
      if (value === null) fail('an invalid hyphen character');
      if (value === '\u001f') run.append(wElement(doc, 'softHyphen'));
      else {
        const text = wElement(doc, 't');
        writeWordText(text, value);
        run.append(text);
      }
    } else if (piece.type === 'hardBreak') run.append(wElement(doc, 'br'));
    else if (piece.type === 'wordTab') run.append(wElement(doc, 'tab'));
    else if (piece.type === 'wordPageBreak') run.append(wElement(doc, 'br', { type: 'page' }));
    else if (piece.type === 'wordColumnBreak') run.append(wElement(doc, 'br', { type: 'column' }));
    else fail('adding this inline object');
    links.append(runs, run, [], piece.marks || []);
  }
  p.append(...runs);
  return p;
}

export async function exportRetainedDocument(
  file: OfficeFile,
  layout?: import('./word-export-layout').WordExportLayout,
): Promise<Blob> {
  if (layout && layout.content !== JSON.stringify(file.content))
    throw Error('Document changed before export. Try exporting again.');
  const { hydrateWordStructure } = await import('./word-structure');
  file = await hydrateWordStructure(file);
  if (file.content.kind !== 'word' || !file.original)
    throw new Error('No retained Word document is available.');
  // Native backups can carry retained bytes that have not passed the Office import path.
  inspectZip(file.original.data);
  const source = await readDocx(file.original.data, {
    legacyRunColors: file.content.runColorsVersion !== 1,
    legacyFontFeatures: file.content.fontFeaturesVersion !== 1,
    legacyParagraphScripts: file.content.paragraphScriptsVersion !== 1,
    legacyHyphens: file.content.hyphenVersion !== 1,
  });
  const zip = await JSZip.loadAsync(file.original.data);
  const addedStoryParts = await prepareWordStoryClones(zip, source, file.content);
  const stories = (value: typeof file.content.stories) => ({
    evenAndOddHeaders: value?.evenAndOddHeaders || false,
    parts: (value?.parts || [])
      .map((part) => ({
        ...part,
        relationshipIds: [...part.relationshipIds].sort(),
        html: wordJSON(sanitizeHTML(part.html)),
      }))
      .sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0)),
  });
  const originalStories = stories(source.content.stories),
    currentStories = stories(file.content.stories);
  const storyMetadata = (value: typeof originalStories) => ({
    parts: value.parts.map(({ html: _html, ...metadata }) => metadata),
  });
  if (!equal(storyMetadata(originalStories), storyMetadata(currentStories)))
    fail('changing header/footer references or page options requires story reference repair');
  if (
    file.content.docxStructure &&
    !equal(
      docxStructureSchema.parse(file.content.docxStructure),
      docxStructureSchema.parse(source.content.docxStructure),
    )
  )
    fail('source-section identities have changed; section restructuring requires reference repair');
  const links = await docxLinks(zip, source.documents);
  const insertedRuns = new Set<Element>();
  const before = wordJSON(sanitizeHTML(source.content.html)),
    after = wordJSON(sanitizeHTML(file.content.html));
  // Freeze requested mark properties before joined source paragraphs are
  // removed or edited. Repeated joins and saved reload retain their origin.
  const markSources: ParagraphMarkSources = new Map();
  const collectMarkSources = (node: JSONContent) => {
    const key = node.attrs?.paragraphMarkSource;
    if (key != null && key !== 'new' && !markSources.has(key)) {
      if (typeof key !== 'string') fail('an invalid paragraph mark source');
      const paragraph = source.paragraphs.get(key);
      if (!paragraph) fail('an unrecognized paragraph mark source; reopen the original document');
      const pr = child(paragraph, 'pPr');
      markSources.set(key, {
        owner: paragraph.ownerDocument,
        properties: (pr && (child(pr, 'rPr')?.cloneNode(true) as Element | undefined)) || null,
      });
    }
    node.content?.forEach(collectMarkSources);
  };
  collectMarkSources(after);
  currentStories.parts.forEach((part) => collectMarkSources(part.html));
  const validateTabContext = (node: JSONContent) => {
    const attrs = node.attrs;
    if (
      attrs?.paragraphDefaultTab != null &&
      attrs.paragraphDefaultTab !== source.tabSettings.interval
    )
      fail('changing the document default tab interval');
    if (
      attrs?.paragraphDecimalSymbol != null &&
      attrs.paragraphDecimalSymbol !== (wordDecimalSymbol(source.tabSettings.decimal) || '.')
    )
      fail('changing the document decimal-tab separator');
    node.content?.forEach(validateTabContext);
  };
  validateTabContext(after);
  currentStories.parts.forEach((part) => validateTabContext(part.html));
  const body = descendants(source.documents.get('word/document.xml')!, 'body')[0];
  // Prepare implicit properties for later section/options edits. Preparation
  // alone must not dirty the body: unchanged exports retain original bytes.
  materializeWordFinalSection(body, source.content.docxStructure?.compatibility?.mode);
  // Finalize page options before the session allocator takes its settings copy.
  // A later body session must not overwrite an odd/even change made after a
  // preceding header session had already claimed its ID.
  const storyOptionsChanged = await writeWordStoryOptions(zip, body, source.content, file.content);
  const textContent = (node: JSONContent): string =>
    node.text ?? (node.content || []).map(textContent).join('');
  const firstSession =
    source.documents.size === 1 && !equal(before, after)
      ? await initializeWordSourceSession(zip, body.ownerDocument, file.content.initialEditSession)
      : undefined;
  const editId =
    firstSession ??
    (textContent(before) === textContent(after) &&
    originalStories.parts.every(
      (part, i) => textContent(part.html) === textContent(currentStories.parts[i].html),
    )
      ? () => undefined
      : await wordEditSession(zip));
  if (file.content.sectionState)
    validateSectionState(
      file.content.sectionState,
      source.content.docxStructure!,
      getSchema(wordExtensions()).nodeFromJSON(after),
    );
  const originalNodes = new Map<string, JSONContent>();
  const oldOrder: string[] = [],
    order: string[] = [];
  const sequence: JSONContent[] = [];
  const parents = new Map<JSONContent, string>();
  const baselineCounts = new Map<string, number>(),
    currentCounts = new Map<string, number>();
  const sectionEndings = new Set(
    source.content.docxStructure?.sections.map((s) => s.endingParagraph).filter(Boolean),
  );
  const liveSections = file.content.sectionState;
  const opaque = new Map<string, string[]>();
  let emptySlots = 0;
  const virtualEmpty = new Set<JSONContent>();
  function walk(node: JSONContent, baseline: boolean, parent = 'doc') {
    if (paragraph(node)) {
      let key = keyOf(node);
      if (baseline) {
        if (!key && !node.content?.length) emptySlots++;
        if (!key && node.content?.length) {
          const identity = JSON.stringify(node);
          key = `opaque:${originalNodes.size}`;
          opaque.set(identity, [...(opaque.get(identity) || []), key]);
          node.attrs = { ...node.attrs, sourceParagraph: key };
        }
        if (key) {
          const occurrence = baselineCounts.get(key) || 0;
          baselineCounts.set(key, occurrence + 1);
          if (occurrence) {
            key = `${key}~${occurrence}`;
            node.attrs = { ...node.attrs, sourceParagraph: key };
          }
          originalNodes.set(key, node);
          oldOrder.push(key);
        }
      } else {
        if (!key && !node.content?.length && emptySlots > 0) {
          emptySlots--;
          virtualEmpty.add(node);
        }
        if (!key) {
          key = opaque.get(JSON.stringify(node))?.shift() || null;
          if (key) node.attrs = { ...node.attrs, sourceParagraph: key };
        }
        if (key) {
          if (!baselineCounts.has(key))
            fail('unrecognized source paragraph; reopen the original document');
          const occurrence = currentCounts.get(key) || 0;
          currentCounts.set(key, occurrence + 1);
          if (!liveSections && occurrence >= baselineCounts.get(key)! && sectionEndings.has(key))
            fail('splitting a section-ending paragraph requires section-break reference repair');
          if (occurrence >= baselineCounts.get(key)!) key = null;
          else if (occurrence) key = `${key}~${occurrence}`;
          node.attrs = { ...node.attrs, sourceParagraph: key };
          if (key) order.push(key);
        }
        sequence.push(node);
        parents.set(node, parent);
      }
    } else node.content?.forEach((n) => walk(n, baseline, node.type || 'doc'));
  }
  walk(before, true);
  walk(after, false);
  const retained = new Set(order);
  if (
    !equal(
      oldOrder.filter((k) => retained.has(k)),
      order,
    )
  )
    fail('reordering source paragraphs');
  const deleted = new Set(oldOrder.filter((k) => !retained.has(k)));
  function skeleton(node: JSONContent, baseline: boolean): unknown {
    if (paragraph(node)) {
      const key = keyOf(node);
      if (key && deleted.has(key)) return null;
      if (key) return { paragraph: key };
      // New paragraphs are handled below. Existing unmapped content must match exactly.
      if (!baseline) return null;
      if (!node.content?.length) return null;
      return node;
    }
    // Flat lists remain a single numbering sequence when a section command
    // splits an item. Compare their retained paragraph order and list attrs;
    // only the newly inserted item wrapper is immaterial to source structure.
    const flatList = ['orderedList', 'bulletList'].includes(node.type || '') &&
      node.content?.every(item => item.type === 'listItem' && item.content?.length === 1 &&
        item.content[0].type === 'paragraph' && !Object.keys(item.attrs || {}).length);
    const children = flatList ? node.content!.map(item => item.content![0]) : node.content || [];
    const content = children
      .map((n) => skeleton(n, baseline))
      .filter((n) => n !== null);
    if (
      ['listItem', 'orderedList', 'bulletList', 'blockquote'].includes(node.type || '') &&
      !content.length
    )
      return null;
    return { ...node, content };
  }
  if (!equal(skeleton(before, true), skeleton(after, false)))
    fail('changing tables, lists, images or unmapped document structures');
  const dirty = new Set<Document>();
  if (firstSession) dirty.add(body.ownerDocument);
  if (storyOptionsChanged.bodyChanged) dirty.add(body.ownerDocument);
  // Story-local paragraph edits share the body's preservation engine. Keep
  // references and opaque structures stable; top-level plain paragraphs can
  // split/join without creating a different header/footer part.
  for (const [index, part] of originalStories.parts.entries()) {
    const next = currentStories.parts[index];
    if (equal(part.html, next.html)) continue;
    const oldParagraphs = part.html.content || [],
      newParagraphs = next.html.content || [];
    const storyDocument = source.documents.get(part.path)!;
    const storyRoot = storyDocument.documentElement;
    const oldKeys = oldParagraphs.map(keyOf);
    const plainStory =
      oldParagraphs.length > 0 &&
      newParagraphs.length > 0 &&
      oldParagraphs.every(paragraph) &&
      newParagraphs.every(paragraph) &&
      oldKeys.every((key) => !!key) &&
      new Set(oldKeys).size === oldKeys.length &&
      storyRoot.children.length === oldParagraphs.length &&
      oldKeys.every((key, i) => source.paragraphs.get(key!) === storyRoot.children[i]);
    if (plainStory) {
      const oldByKey = new Map(oldParagraphs.map((node) => [keyOf(node)!, node]));
      const retained = new Set<string>();
      const sequence = newParagraphs.map((node) => {
        const key = keyOf(node);
        if (key && !oldByKey.has(key))
          fail('header/footer paragraph identities have changed or are ambiguous');
        // Enter duplicates the source attribute. Its first occurrence keeps
        // the source XML; subsequent occurrences are newly authored paragraphs.
        if (!key || retained.has(key)) return { node, key: null };
        retained.add(key);
        return { node, key };
      });
      if (
        !equal(
          oldKeys.filter((key) => retained.has(key!)),
          sequence.map((n) => n.key).filter(Boolean),
        )
      )
        fail('reordering header/footer source paragraphs');
      for (const key of oldKeys) {
        if (retained.has(key!)) continue;
        const p = source.paragraphs.get(key!)!;
        // Deleting text must never silently delete fields, drawings, bookmarks,
        // list references or other structures not represented by this editor.
        const unknown = paragraphDeletionStructures(p);
        if (unknown.length)
          fail(
            `deleting a header/footer paragraph containing document structures (${[...new Set(unknown.map((e) => e.localName))].slice(0, 8).join(', ')})`,
          );
      }
      const output: Element[] = [];
      for (const [index, { node, key }] of sequence.entries()) {
        if (key) {
          const p = source.paragraphs.get(key)!;
          let old = oldByKey.get(key)!;
          const nextText = flat(node);
          // Joining contiguous source paragraphs moves their existing runs.
          // Preserve that origin, including paragraph-inherited run sessions,
          // instead of treating the moved suffix as newly typed text. Story
          // typing carries wordEditRun, so retyping an identical suffix cannot
          // accidentally masquerade as a source join.
          for (let at = oldKeys.indexOf(key) + 1; at < oldKeys.length; at++) {
            const movedKey = oldKeys[at]!;
            if (retained.has(movedKey)) break;
            const moved = oldByKey.get(movedKey)!;
            const prefix = flat(old),
              suffix = flat(moved);
            if (
              !suffix.text ||
              !nextText.text.startsWith(prefix.text + suffix.text) ||
              nextText.marks
                .slice(prefix.text.length, prefix.text.length + suffix.text.length)
                .some((marks) => wordRunSession(marks))
            )
              break;
            const sourceParagraph = source.paragraphs.get(movedKey)!;
            const inheritedSession = val(sourceParagraph, 'rsidRDefault');
            const runElements = [...sourceParagraph.children].filter((e) => e.localName !== 'pPr');
            if (
              runElements.some(
                (e) => e.localName !== 'r' || (!val(e, 'rsidR') && !inheritedSession),
              )
            )
              fail('joining header/footer paragraphs without retained editing provenance');
            for (const sourceRun of runElements) {
              const run = sourceRun.cloneNode(true) as Element;
              if (
                !val(run, 'rsidR') &&
                inheritedSession &&
                inheritedSession !== val(p, 'rsidRDefault')
              )
                setVal(run, inheritedSession, 'rsidR');
              p.append(run);
            }
            old = { ...old, content: [...(old.content || []), ...(moved.content || [])] };
          }
          if (!equal(old, node))
            patchParagraph(p, old, node, links, insertedRuns, editId, markSources);
          output.push(p);
        } else {
          const template = output.at(-1) || storyRoot.firstElementChild!;
          const added = newParagraph(
            storyDocument,
            node,
            links,
            editId,
            markSources,
            template,
            sequence[index - 1]?.node || oldParagraphs[0],
          );
          const originKey = keyOf(node),
            origin = originKey && source.paragraphs.get(originKey);
          const defaultSession = (origin && val(origin, 'rsidRDefault')) || editId();
          if (defaultSession) setVal(added, defaultSession, 'rsidRDefault');
          output.push(added);
        }
      }
      for (const key of oldKeys) if (!retained.has(key!)) source.paragraphs.get(key!)!.remove();
      storyRoot.append(...output);
      dirty.add(storyDocument);
      continue;
    }
    const edits: { before: JSONContent; after: JSONContent; source: Element }[] = [];
    const seen = new Set<string>();
    const visit = (old: JSONContent, current: JSONContent) => {
      if (paragraph(old) && paragraph(current)) {
        const key = keyOf(old),
          p = key && source.paragraphs.get(key);
        if (
          !key ||
          key !== keyOf(current) ||
          seen.has(key) ||
          !p ||
          p.ownerDocument !== source.documents.get(part.path)
        )
          fail('header/footer paragraph identities have changed or are ambiguous');
        seen.add(key);
        if (!equal(old, current)) edits.push({ before: old, after: current, source: p });
        return;
      }
      const { content: oldChildren = [], ...oldShape } = old;
      const { content: newChildren = [], ...newShape } = current;
      if (!equal(oldShape, newShape) || oldChildren.length !== newChildren.length)
        fail('header/footer structural editing requires reference repair');
      oldChildren.forEach((child, i) => visit(child, newChildren[i]));
    };
    visit(part.html, next.html);
    for (const edit of edits) {
      patchParagraph(
        edit.source,
        edit.before,
        edit.after,
        links,
        insertedRuns,
        editId,
        markSources,
      );
      dirty.add(edit.source.ownerDocument);
    }
  }
  const commitSections = file.content.sectionState
    ? prepareLiveSections(body, source.content.docxStructure!, file.content.sectionState)
    : null;
  for (const key of deleted) {
    const p = source.paragraphs.get(key);
    if (!p) fail('deleting an unmapped source paragraph');
    if (paragraphDeletionStructures(p).length)
      fail('deleting a paragraph containing document structures');
    // A table cell must retain a final paragraph.
    if (p.parentElement?.localName !== 'body') fail('deleting paragraphs inside tables or notes');
    dirty.add(p.ownerDocument);
    p.remove();
  }
  for (const node of sequence) {
    const key = keyOf(node);
    if (!key) continue;
    const old = originalNodes.get(key)!;
    if (!equal(old, node)) {
      const p = source.paragraphs.get(key);
      if (!p) fail('editing an unmapped source paragraph');
      patchParagraph(p, old, node, links, insertedRuns, editId, markSources);
      dirty.add(p.ownerDocument);
    }
  }
  const outputParagraphs = new Map<JSONContent, Element>();
  let previous: Element | undefined;
  for (let i = 0; i < sequence.length; i++) {
    const node = sequence[i],
      key = keyOf(node);
    if (key) {
      previous = source.paragraphs.get(key) || previous;
      if (previous) outputParagraphs.set(node, previous);
      continue;
    }
    if (virtualEmpty.has(node)) continue;
    const nextKey = sequence
      .slice(i + 1)
      .map(keyOf)
      .find(Boolean);
    const next = nextKey ? source.paragraphs.get(nextKey) : undefined;
    let parent: Element = body,
      anchor: Element | null = child(body, 'sectPr') || null;
    if (parents.get(node) === 'doc') {
      const top = (p: Element | undefined) => {
        while (p?.parentElement && p.parentElement !== body) p = p.parentElement;
        return p?.parentElement === body ? p : undefined;
      };
      const previousBlock = top(previous),
        nextBlock = top(next);
      anchor = previousBlock ? previousBlock.nextElementSibling : nextBlock || anchor;
    } else if (previous && (!next || previous.parentElement === next.parentElement)) {
      parent = previous.parentElement!;
      anchor = previous.nextElementSibling;
    } else if (next && (!previous || next.parentElement === body)) {
      parent = next.parentElement!;
      anchor = next;
    } else if (previous) fail('inserting a paragraph between different document structures');
    const previousProperties = previous && child(previous, 'pPr');
    if (
      parents.get(node) === 'listItem' &&
      !previous?.parentElement?.localName.match(/^(footnote|endnote)$/) &&
      !(previousProperties && child(previousProperties, 'numPr'))
    )
      fail('adding a new list without a mapped numbering definition');
    const added = newParagraph(
      parent.ownerDocument,
      node,
      links,
      editId,
      markSources,
      parents.get(node) === 'listItem' ? previous : undefined,
      parents.get(node) === 'listItem' ? sequence[i - 1] : undefined,
    );
    parent.insertBefore(added, anchor);
    previous = added;
    outputParagraphs.set(node, added);
    dirty.add(parent.ownerDocument);
  }
  if (commitSections?.(sequence.map((node) => outputParagraphs.get(node))))
    dirty.add(body.ownerDocument);
  if (
    file.content.sectionState?.version === 2 &&
    (await writeWordStoryOptions(zip, body, source.content, file.content, true)).bodyChanged
  )
    dirty.add(body.ownerDocument);
  if (writeWordStoryReferences(body, file.content)) dirty.add(body.ownerDocument);
  if (
    file.content.pageOverrides?.paper ||
    file.content.pageOverrides?.orientation ||
    file.content.pageOverrides?.margin ||
    file.content.paper !== source.content.paper ||
    file.content.margin !== source.content.margin ||
    file.content.orientation !== source.content.orientation
  ) {
    for (const section of liveSectionProperties(body)) {
      if (
        file.content.pageOverrides?.paper ||
        file.content.pageOverrides?.orientation ||
        file.content.paper !== source.content.paper ||
        file.content.orientation !== source.content.orientation
      ) {
        const size = property(section, 'pgSz');
        let dimensions = [Number(val(size, 'w')), Number(val(size, 'h'))];
        if (
          file.content.pageOverrides?.paper ||
          file.content.paper !== source.content.paper ||
          dimensions.some((n) => !Number.isFinite(n) || n <= 0)
        )
          dimensions = file.content.paper === 'a4' ? [11906, 16838] : [12240, 15840];
        dimensions.sort((a, b) => a - b);
        if (file.content.orientation === 'landscape') dimensions.reverse();
        setVal(size, String(dimensions[0]), 'w');
        setVal(size, String(dimensions[1]), 'h');
        setVal(size, file.content.orientation || 'portrait', 'orient');
      }
      if (file.content.pageOverrides?.margin || file.content.margin !== source.content.margin) {
        const margin = property(section, 'pgMar');
        for (const side of ['left', 'right', 'top', 'bottom'])
          setVal(
            margin,
            file.content.margin === 'narrow'
              ? '720'
              : file.content.margin === 'wide'
                ? '2160'
                : '1440',
            side,
          );
      }
    }
    dirty.add(body.ownerDocument);
  }
  if (file.content.sectionState?.version === 2 && file.content.sectionState.layouts?.length) {
    const { writeWordSectionGeometry } = await import('./docx-section-geometry');
    if (writeWordSectionGeometry(body, file.content)) dirty.add(body.ownerDocument);
  }
  if (!dirty.size && !storyOptionsChanged.settingsChanged && !addedStoryParts)
    return new Blob([file.original.data], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
  links.save();
  for (const [path, doc] of source.documents) {
    if (!dirty.has(doc)) continue;
    if (path === 'word/document.xml') {
      coalescePlainWordRuns(doc, insertedRuns);
      normalizeWordTextSpace(doc);
      if (layout) {
        const { splitWordColumnRuns } = await import('./docx-column-runs');
        splitWordColumnRuns(doc, layout);
      }
    }
    zip.file(path, serial(doc), { createFolders: false });
  }
  await registerMissingWordSessions(zip, source.documents.values());
  return zip.generateAsync({
    type: 'blob',
    compression: 'DEFLATE',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}
