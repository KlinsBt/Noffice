import JSZip from 'jszip';
import { localGroupGeometry } from './pptx-group-transform';
import { normalizeGroupBounds } from './pptx-group-bounds';
import { textElement, type OfficeFile, type SlideElement } from './model';
import { contentFingerprint } from './office-preservation';
import { elements, parseXML } from './xlsx-import';
import { readPresentation, presentationRelations } from './pptx-import';
import { pptxStackNodes } from './pptx-stack';
import { stackKey } from './slide-stack';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE = 'http://schemas.openxmlformats.org/package/2006/relationships';
async function addRelationship(zip: JSZip, from: string, to: string, type: string) {
  const path = from.replace(/([^/]+)$/, '_rels/$1.rels');
  const existing = await zip.file(path)?.async('string');
  const doc = parseXML(existing || `<Relationships xmlns="${PACKAGE}"/>`),
    root = doc.documentElement;
  const ids = new Set(elements(doc, 'Relationship').map((e) => e.getAttribute('Id')));
  let id = 1;
  while (ids.has('rId' + id)) id++;
  const base = from.split('/').slice(0, -1),
    target = to.split('/');
  while (base.length && target.length && base[0] === target[0]) {
    base.shift();
    target.shift();
  }
  const rel = doc.createElementNS(PACKAGE, 'Relationship');
  rel.setAttribute('Id', 'rId' + id);
  rel.setAttribute('Type', R + '/' + type);
  rel.setAttribute('Target', '../'.repeat(base.length) + target.join('/'));
  root.appendChild(rel);
  zip.file(path, new XMLSerializer().serializeToString(doc), { createFolders: false });
  return 'rId' + id;
}
async function registerPart(zip: JSZip, path: string, mime: string) {
  const doc = parseXML(await zip.file('[Content_Types].xml')!.async('string'));
  const entry = doc.createElementNS(doc.documentElement.namespaceURI, 'Override');
  entry.setAttribute('PartName', '/' + path);
  entry.setAttribute('ContentType', mime);
  doc.documentElement.appendChild(entry);
  zip.file('[Content_Types].xml', new XMLSerializer().serializeToString(doc), {
    createFolders: false,
  });
}
function availablePart(zip: JSZip, prefix: string, extension: string) {
  let n = 1;
  while (zip.file(prefix + n + '.' + extension)) n++;
  return prefix + n + '.' + extension;
}
async function makePicture(
  zip: JSZip,
  doc: Document,
  path: string,
  item: SlideElement,
  id: number,
  width: number,
  height: number,
) {
  const match = /^data:image\/(png|jpeg|gif|webp);base64,([A-Za-z0-9+/]+={0,2})$/.exec(
    item.src || '',
  );
  if (!match || match[2].length > Math.ceil((8 * 1024 * 1024) / 3) * 4)
    throw Error('Choose a supported local image smaller than 8 MB.');
  const media = availablePart(zip, 'ppt/media/noffice-image-', match[1]);
  zip.file(media, match[2], { base64: true, createFolders: false });
  await registerPart(zip, media, 'image/' + match[1]);
  const rel = await addRelationship(zip, path, media, 'image');
  const pic = doc.createElementNS(P, 'p:pic'),
    nv = doc.createElementNS(P, 'p:nvPicPr');
  pic.appendChild(nv);
  const props = doc.createElementNS(P, 'p:cNvPr');
  props.setAttribute('id', String(id));
  props.setAttribute('name', 'Noffice image ' + id);
  nv.appendChild(props);
  nv.appendChild(doc.createElementNS(P, 'p:cNvPicPr'));
  nv.appendChild(doc.createElementNS(P, 'p:nvPr'));
  const fill = doc.createElementNS(P, 'p:blipFill');
  pic.appendChild(fill);
  ensureDrawing(fill, 'blip').setAttributeNS(R, 'r:embed', rel);
  ensureDrawing(ensureDrawing(fill, 'stretch'), 'fillRect');
  const generated = makeShape(
    doc,
    { ...item, type: 'rect', text: '', fill: 'transparent' },
    id,
    width,
    height,
  );
  const geometry = direct(generated, 'spPr')!;
  const transform = ensureDrawing(geometry, 'xfrm');
  if (item.flipH) transform.setAttribute('flipH', '1');
  if (item.flipV) transform.setAttribute('flipV', '1');
  pic.appendChild(geometry);
  return pic;
}
async function ensureNotesMaster(zip: JSZip) {
  const relations = await presentationRelations(zip, 'ppt/presentation.xml');
  const existing = [...relations.values()].find((p) => /notesMasters\/[^/]+\.xml$/.test(p));
  if (existing) {
    if (!zip.file(existing)) throw Error('The presentation references a missing notes master.');
    return existing;
  }
  const path = availablePart(zip, 'ppt/notesMasters/notesMaster', 'xml');
  const master = `<p:notesMaster xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr><a:defRPr sz="1200"><a:latin typeface="Arial"/></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>`;
  zip.file(path, master, { createFolders: false });
  await registerPart(
    zip,
    path,
    'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
  );
  const theme = [...relations.values()].find((p) => /\/theme\//.test(p));
  if (theme) {
    // The pinned PowerPoint build rejects a notes master sharing the slide theme part.
    // Retain the theme payload and any relative asset relationships in the same directory.
    const data = await zip.file(theme)?.async('uint8array');
    if (!data) throw Error('The presentation references a missing theme.');
    const notesTheme = availablePart(zip, theme.replace(/[^/]+$/, 'noffice-notes-theme'), 'xml');
    zip.file(notesTheme, data, { createFolders: false });
    const themeRels = await zip
      .file(theme.replace(/([^/]+)$/, '_rels/$1.rels'))
      ?.async('uint8array');
    if (themeRels)
      zip.file(notesTheme.replace(/([^/]+)$/, '_rels/$1.rels'), themeRels, {
        createFolders: false,
      });
    await registerPart(zip, notesTheme, 'application/vnd.openxmlformats-officedocument.theme+xml');
    await addRelationship(zip, path, notesTheme, 'theme');
  }
  const id = await addRelationship(zip, 'ppt/presentation.xml', path, 'notesMaster');
  const doc = parseXML(await zip.file('ppt/presentation.xml')!.async('string'));
  const list = doc.createElementNS(P, 'p:notesMasterIdLst');
  const entry = doc.createElementNS(P, 'p:notesMasterId');
  entry.setAttributeNS(R, 'r:id', id);
  list.appendChild(entry);
  // Presentation order: slide masters, notes masters, handout masters, slide IDs, sizes.
  doc.documentElement.insertBefore(
    list,
    [...doc.documentElement.children].find((e) => e.localName !== 'sldMasterIdLst') || null,
  );
  zip.file('ppt/presentation.xml', new XMLSerializer().serializeToString(doc), {
    createFolders: false,
  });
  return path;
}
async function addNotes(zip: JSZip, slidePath: string, text: string) {
  const path = availablePart(zip, 'ppt/notesSlides/notesSlide', 'xml');
  const doc = parseXML(
    `<p:notes xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`,
  );
  const shape = makeShape(
    doc,
    textElement(text, { x: 60, y: 270, w: 840, h: 220, fontSize: 14, color: '#000000' }),
    2,
    6858000,
    9144000,
  );
  const placeholder = doc.createElementNS(P, 'p:ph');
  placeholder.setAttribute('type', 'body');
  placeholder.setAttribute('idx', '1');
  elements(shape, 'nvPr')[0].appendChild(placeholder);
  elements(doc, 'spTree')[0].appendChild(shape);
  await addRelationship(zip, slidePath, path, 'notesSlide');
  await addRelationship(zip, path, slidePath, 'slide');
  const master = await ensureNotesMaster(zip);
  await addRelationship(zip, path, master, 'notesMaster');
  zip.file(path, new XMLSerializer().serializeToString(doc), { createFolders: false });
  await registerPart(
    zip,
    path,
    'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
  );
}

const direct = (parent: Element, name: string) =>
  Array.from(parent.children).find((e) => e.localName === name);
function ensureDrawing(parent: Element, name: string) {
  let element = direct(parent, name);
  if (!element) {
    element = parent.ownerDocument.createElementNS(A, 'a:' + name);
    const orders: Record<string, string[]> = {
      spPr: [
        'xfrm',
        'prstGeom',
        'custGeom',
        'noFill',
        'solidFill',
        'gradFill',
        'blipFill',
        'pattFill',
        'grpFill',
        'ln',
        'effectLst',
        'effectDag',
        'scene3d',
        'sp3d',
        'extLst',
      ],
      rPr: [
        'ln',
        'noFill',
        'solidFill',
        'gradFill',
        'blipFill',
        'pattFill',
        'grpFill',
        'effectLst',
        'effectDag',
        'highlight',
        'uLnTx',
        'uLn',
        'uFillTx',
        'uFill',
        'latin',
        'ea',
        'cs',
        'sym',
        'hlinkClick',
        'hlinkMouseOver',
        'rtl',
        'extLst',
      ],
      p: ['pPr', 'r', 'br', 'fld', 'endParaRPr'],
      ln: [
        'noFill',
        'solidFill',
        'gradFill',
        'pattFill',
        'prstDash',
        'custDash',
        'round',
        'bevel',
        'miter',
        'headEnd',
        'tailEnd',
        'extLst',
      ],
      xfrm: ['off', 'ext', 'chOff', 'chExt'],
    };
    const order =
      orders[['endParaRPr', 'defRPr'].includes(parent.localName) ? 'rPr' : parent.localName];
    parent.insertBefore(
      element,
      order
        ? Array.from(parent.children).find(
            (e) => order.indexOf(e.localName) > order.indexOf(name),
          ) || null
        : null,
    );
  }
  return element;
}
function replaceFill(parent: Element, value: string, opacity = 1) {
  for (const child of Array.from(parent.children))
    if (
      ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'].includes(
        child.localName,
      )
    )
      child.remove();
  if (value === 'transparent') ensureDrawing(parent, 'noFill');
  else {
    if (!/^#[a-f\d]{6}$/i.test(value)) throw Error('Choose a supported six-digit slide color.');
    ensureDrawing(ensureDrawing(parent, 'solidFill'), 'srgbClr').setAttribute(
      'val',
      value.slice(1),
    );
    if (opacity < 1)
      ensureDrawing(
        ensureDrawing(ensureDrawing(parent, 'solidFill'), 'srgbClr'),
        'alpha',
      ).setAttribute('val', String(Math.round(opacity * 100000)));
  }
}
function patchOutline(
  properties: Element,
  item: SlideElement,
  width: number,
  previous?: SlideElement,
) {
  const outline = item.outline;
  if (!outline) return;
  const line = ensureDrawing(properties, 'ln');
  const before = previous?.outline;
  if (!before || outline.width !== before.width)
    line.setAttribute('w', String(Math.round((outline.width * width) / 960)));
  if (!before || outline.color !== before.color || outline.opacity !== before.opacity)
    replaceFill(line, outline.color, outline.opacity);
  if (!before || outline.dash !== before.dash) {
    direct(line, 'custDash')?.remove();
    ensureDrawing(line, 'prstDash').setAttribute('val', outline.dash);
  }
}
function patchText(body: Element, text: string) {
  if (elements(body, 'fld').length || elements(body, 'br').length)
    throw Error('Editing text containing slide fields or soft line breaks is not supported yet.');
  const paragraphs = Array.from(body.children).filter((e) => e.localName === 'p'),
    lines = text.split('\n');
  const template = paragraphs.at(-1);
  for (let i = 0; i < lines.length; i++) {
    let paragraph = paragraphs[i];
    if (!paragraph) {
      paragraph = body.ownerDocument.createElementNS(A, 'a:p');
      const properties = template && direct(template, 'pPr');
      if (properties) paragraph.appendChild(properties.cloneNode(true));
      body.appendChild(paragraph);
    }
    let texts = elements(paragraph, 't');
    if (!texts.length) {
      const run = ensureDrawing(paragraph, 'r');
      const format = template && elements(template, 'rPr')[0];
      if (format) run.appendChild(format.cloneNode(true));
      texts = [ensureDrawing(run, 't')];
    }
    const previous = texts.map((t) => t.textContent || '').join(''),
      next = lines[i];
    if (previous === next) continue;
    let prefix = 0,
      suffix = 0;
    while (prefix < previous.length && prefix < next.length && previous[prefix] === next[prefix])
      prefix++;
    while (
      suffix < previous.length - prefix &&
      suffix < next.length - prefix &&
      previous[previous.length - 1 - suffix] === next[next.length - 1 - suffix]
    )
      suffix++;
    const until = previous.length - suffix,
      replacement = next.slice(prefix, next.length - suffix);
    let offset = 0,
      inserted = false;
    for (const node of texts) {
      const raw = node.textContent || '',
        end = offset + raw.length;
      const insert =
        !inserted &&
        (prefix < end || (prefix === end && (prefix === until || node === texts.at(-1))));
      node.textContent =
        raw.slice(0, Math.max(0, prefix - offset)) +
        (insert ? replacement : '') +
        raw.slice(Math.max(0, until - offset));
      if (insert) inserted = true;
      offset = end;
    }
  }
  paragraphs.slice(lines.length).forEach((p) => p.remove());
}

/** Patch supported existing-object edits into their original parts. */
export async function exportRetainedPresentation(file: OfficeFile): Promise<Blob> {
  const { hydratePresentationText } = await import('./presentation-migration');
  file = await hydratePresentationText(file);
  if (file.content.kind !== 'powerpoint' || !file.original)
    throw Error('Missing retained presentation.');
  const content = file.content,
    baseline = await readPresentation(file.original.data),
    zip = await JSZip.loadAsync(file.original.data);
  if (
    content.aspectRatio !== undefined &&
    Math.abs(content.aspectRatio - baseline.aspectRatio!) > 0.000001
  )
    throw Error('Changing imported slide dimensions requires layout-aware editing.');
  const sourceSlides = new Map(baseline.slides.map((s) => [s.sourcePath, s]));
  if (
    content.slides.length !== baseline.slides.length ||
    new Set(content.slides.map((s) => s.sourcePath)).size !== baseline.slides.length ||
    content.slides.some((s) => !sourceSlides.has(s.sourcePath))
  )
    throw Error(
      'Imported slide insertion, deletion and duplication require relationship-aware editing. Keep the original slides or export a Noffice backup.',
    );
  const presentation = parseXML(await zip.file('ppt/presentation.xml')!.async('string')),
    size = elements(presentation, 'sldSz')[0];
  const width = Number(size?.getAttribute('cx') || 12192000),
    height = Number(size?.getAttribute('cy') || 6858000);
  if (content.slides.some((s, i) => s.sourcePath !== baseline.slides[i].sourcePath)) {
    if (elements(presentation, 'sectionLst').length)
      throw Error('Reordering a presentation with sections requires section-aware editing.');
    const list = direct(presentation.documentElement, 'sldIdLst');
    if (!list) throw Error('The presentation has no slide order.');
    const relations = await presentationRelations(zip, 'ppt/presentation.xml');
    const ids = new Map(
      Array.from(list.children)
        .filter((e) => e.localName === 'sldId')
        .map((e) => [relations.get(e.getAttributeNS(R, 'id') || e.getAttribute('r:id') || ''), e]),
    );
    for (const slide of content.slides) {
      const id = ids.get(slide.sourcePath);
      if (!id) throw Error('A source slide relationship is missing.');
      list.appendChild(id);
    }
    zip.file('ppt/presentation.xml', new XMLSerializer().serializeToString(presentation), {
      createFolders: false,
    });
  }
  for (let i = 0; i < content.slides.length; i++) {
    const slide = content.slides[i],
      old = sourceSlides.get(slide.sourcePath)!;
    const imported = slide.elements.filter((e) => e.sourceShapeId);
    const oldObjects = new Map(old.elements.map((e) => [e.sourceShapeId, e]));
    if (new Set(imported.map((e) => e.sourceShapeId)).size !== imported.length)
      throw Error('Imported object identities must be unique.');
    if (
      imported.length !== old.elements.length ||
      imported.some((e) => !oldObjects.has(e.sourceShapeId))
    )
      throw Error(
        'Removing or duplicating imported objects is not supported by the preservation writer yet.',
      );
    const doc = parseXML(await zip.file(slide.sourcePath!)!.async('string')),
      tree = elements(doc, 'spTree')[0];
    if (!tree) throw Error('The imported slide has no shape tree.');
    const stackNodes = pptxStackNodes(tree);
    const changedShapes: Element[] = [];
    let dirty = false;
    for (let j = 0; j < imported.length; j++) {
      const item = imported[j],
        previous = oldObjects.get(item.sourceShapeId)!;
      const changed = (key: keyof SlideElement) =>
        ['fill', 'color'].includes(key)
          ? String(item[key]).toLowerCase() !== String(previous[key]).toLowerCase()
          : item[key] !== previous[key];
      if (item.type !== previous.type || item.src !== previous.src)
        throw Error('Replacing imported object types or image data is not supported yet.');
      const shape = [...elements(doc, 'sp'), ...elements(doc, 'pic')].find(
        (e) => elements(e, 'cNvPr')[0]?.getAttribute('id') === item.sourceShapeId,
      );
      if (!shape) throw Error('The original slide object is missing.');
      const rotationChanged =
        item.rotation !== undefined && item.rotation !== (previous.rotation || 0);
      const flips = (['flipH', 'flipV'] as const).filter(
        (key) => item[key] !== undefined && item[key] !== !!previous[key],
      );
      if (flips.length && item.type !== 'image')
        throw Error('Flipping imported text and shapes is not supported yet.');
      const geometry =
        (['x', 'y', 'w', 'h'] as const).some(changed) || rotationChanged || flips.length;
      if (geometry) {
        if (shape.parentElement?.localName !== 'spTree' && !previous.sourceGroupTransform)
          throw Error('Moving or resizing grouped objects requires group-coordinate editing.');
        const local = localGroupGeometry(previous, item, width, height);
        changedShapes.push(shape);
        const properties = direct(shape, 'spPr');
        if (!properties) throw Error('The object has no editable geometry.');
        const transform = ensureDrawing(properties, 'xfrm');
        if (rotationChanged)
          transform.setAttribute('rot', String(Math.round((item.rotation || 0) * 60000)));
        for (const key of flips) transform.setAttribute(key, item[key] ? '1' : '0');
        properties.insertBefore(transform, properties.firstChild);
        const offset = ensureDrawing(transform, 'off'),
          extent = ensureDrawing(transform, 'ext');
        if (changed('x') || previous.sourceGroupTransform)
          offset.setAttribute('x', String(Math.round(local.x)));
        if (changed('y') || previous.sourceGroupTransform)
          offset.setAttribute('y', String(Math.round(local.y)));
        if (changed('w') || previous.sourceGroupTransform)
          extent.setAttribute('cx', String(Math.round(local.w)));
        if (changed('h') || previous.sourceGroupTransform)
          extent.setAttribute('cy', String(Math.round(local.h)));
        // A newly materialized inherited transform requires all coordinates.
        if (!offset.hasAttribute('x')) offset.setAttribute('x', String(Math.round(local.x)));
        if (!offset.hasAttribute('y')) offset.setAttribute('y', String(Math.round(local.y)));
        if (!extent.hasAttribute('cx')) extent.setAttribute('cx', String(Math.round(local.w)));
        if (!extent.hasAttribute('cy')) extent.setAttribute('cy', String(Math.round(local.h)));
        dirty = true;
      }
      const body = direct(shape, 'txBody');
      if (changed('text')) {
        if (!body) throw Error('This object has no editable text body.');
        patchText(body, item.text);
        dirty = true;
      }
      const styles = (
        ['fontSize', 'fontFamily', 'bold', 'italic', 'underline', 'color', 'align'] as const
      ).filter(changed);
      if (styles.length) {
        if (!body) throw Error('This object has no editable text formatting.');
        for (const paragraph of Array.from(body.children).filter((e) => e.localName === 'p')) {
          if (changed('align')) {
            const properties = ensureDrawing(paragraph, 'pPr');
            paragraph.insertBefore(properties, paragraph.firstChild);
            properties.setAttribute('algn', { left: 'l', center: 'ctr', right: 'r' }[item.align]);
          }
          const runs = Array.from(paragraph.children).filter((e) =>
            ['r', 'fld', 'br'].includes(e.localName),
          );
          const formats = runs.map((run) => {
            const properties = ensureDrawing(run, 'rPr');
            run.insertBefore(properties, run.firstChild);
            return properties;
          });
          formats.push(ensureDrawing(paragraph, 'endParaRPr'));
          for (const format of formats) {
            if (changed('fontSize'))
              format.setAttribute('sz', String(Math.round((item.fontSize * width) / 960 / 127)));
            if (changed('fontFamily'))
              ensureDrawing(format, 'latin').setAttribute('typeface', item.fontFamily || 'Arial');
            if (changed('bold')) format.setAttribute('b', item.bold ? '1' : '0');
            if (changed('italic')) format.setAttribute('i', item.italic ? '1' : '0');
            if (changed('underline')) format.setAttribute('u', item.underline ? 'sng' : 'none');
            if (changed('color')) replaceFill(format, item.color);
          }
        }
        dirty = true;
      }
      if (
        changed('fill') ||
        (item.fillOpacity !== undefined && item.fillOpacity !== (previous.fillOpacity ?? 1))
      ) {
        const properties = direct(shape, 'spPr');
        if (!properties || item.type === 'image')
          throw Error('This object fill cannot be edited yet.');
        replaceFill(properties, item.fill, item.fillOpacity ?? previous.fillOpacity ?? 1);
        dirty = true;
      }
      if (item.outline && JSON.stringify(item.outline) !== JSON.stringify(previous.outline)) {
        const properties = direct(shape, 'spPr');
        if (!properties) throw Error('This object outline cannot be edited yet.');
        patchOutline(properties, item, width, previous);
        dirty = true;
      }
    }
    normalizeGroupBounds(changedShapes, width, height);
    const additions = slide.elements.filter((e) => !e.sourceShapeId);
    let id = Math.max(0, ...elements(doc, 'cNvPr').map((e) => Number(e.getAttribute('id')) || 0));
    for (const item of additions) {
      const shape =
        item.type === 'image'
          ? await makePicture(zip, doc, slide.sourcePath!, item, ++id, width, height)
          : makeShape(doc, item, ++id, width, height);
      tree.insertBefore(shape, direct(tree, 'extLst') || null);
      stackNodes.set(stackKey(item), shape);
      dirty = true;
    }
    const additionKeys = new Set(additions.map(stackKey));
    const requested = (slide.stackOrder || [...stackNodes.keys()]).filter(
      (key) => !key.startsWith('new:') || additionKeys.has(key),
    );
    for (const key of additionKeys) if (!requested.includes(key)) requested.push(key);
    if (
      requested.length !== stackNodes.size ||
      new Set(requested).size !== requested.length ||
      requested.some((key) => !stackNodes.has(key))
    )
      throw Error('The imported object stacking order is incomplete or invalid.');
    if (slide.stackOrder) {
      const modeled = [...new Set(slide.elements.map(stackKey))];
      const visible = requested.filter((key) => modeled.includes(key));
      if (visible.length !== modeled.length || visible.some((key, index) => key !== modeled[index]))
        throw Error('The object order and its source stacking metadata disagree.');
    }
    if (requested.some((key, index) => key !== [...stackNodes.keys()][index])) {
      for (const key of requested)
        tree.insertBefore(stackNodes.get(key)!, direct(tree, 'extLst') || null);
      dirty = true;
    }
    if (
      slide.background.toLowerCase() !== old.background.toLowerCase() ||
      (slide.backgroundOpacity !== undefined &&
        slide.backgroundOpacity !== (old.backgroundOpacity ?? 1))
    ) {
      const common = direct(doc.documentElement, 'cSld');
      if (!common) throw Error('Missing slide content.');
      direct(common, 'bg')?.remove();
      const bg = doc.createElementNS(P, 'p:bg'),
        properties = doc.createElementNS(P, 'p:bgPr');
      bg.appendChild(properties);
      replaceFill(
        properties,
        slide.background,
        slide.backgroundOpacity ?? old.backgroundOpacity ?? 1,
      );
      common.insertBefore(bg, common.firstChild);
      dirty = true;
    }
    if (slide.notes !== old.notes) {
      const rels = await presentationRelations(zip, slide.sourcePath!);
      const path = [...rels.values()].find((p) => /notesSlides\/notesSlide\d+\.xml/.test(p));
      if (!path) {
        await addNotes(zip, slide.sourcePath!, slide.notes);
        if (dirty)
          zip.file(slide.sourcePath!, new XMLSerializer().serializeToString(doc), {
            createFolders: false,
          });
        continue;
      }
      const notes = parseXML(await zip.file(path)!.async('string'));
      const bodies = elements(notes, 'sp')
        .filter((sp) => elements(sp, 'ph')[0]?.getAttribute('type') === 'body')
        .map((sp) => direct(sp, 'txBody'))
        .filter((v): v is Element => !!v);
      if (bodies.length !== 1)
        throw Error('This notes page requires a single editable body placeholder.');
      patchText(bodies[0], slide.notes);
      zip.file(path, new XMLSerializer().serializeToString(notes), { createFolders: false });
    }
    if (dirty)
      zip.file(slide.sourcePath!, new XMLSerializer().serializeToString(doc), {
        createFolders: false,
      });
  }
  return new Blob([await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })], {
    type: MIME,
  });
}

/** Preserve all source objects when the only model changes are added text/rectangles/ellipses.
 * Other edits return null to the explicitly partial general writer; they are never discarded.
 */
export async function exportAddedPresentationObjects(file: OfficeFile): Promise<Blob | null> {
  if (
    file.content.kind !== 'powerpoint' ||
    !file.original?.contentFingerprint ||
    !file.content.slides.every((s) => s.sourcePath)
  )
    return null;
  const content = file.content;
  const baseline = {
    ...content,
    slides: content.slides.map((s) => ({
      ...s,
      elements: s.elements.filter((e) => e.sourceShapeId),
    })),
  };
  if ((await contentFingerprint(baseline)) !== file.original.contentFingerprint) return null;
  const added = content.slides.flatMap((s) => s.elements.filter((e) => !e.sourceShapeId));
  if (!added.length || added.some((e) => !['text', 'rect', 'ellipse'].includes(e.type)))
    return null;
  const zip = await JSZip.loadAsync(file.original.data);
  const presentation = parseXML(await zip.file('ppt/presentation.xml')!.async('string'));
  const size = elements(presentation, 'sldSz')[0];
  const width = Number(size?.getAttribute('cx') || 12192000),
    height = Number(size?.getAttribute('cy') || 6858000);
  for (const slide of content.slides) {
    const additions = slide.elements.filter((e) => !e.sourceShapeId);
    if (!additions.length) continue;
    const doc = parseXML(await zip.file(slide.sourcePath!)!.async('string'));
    const tree = elements(doc, 'spTree')[0];
    if (!tree) throw Error('The imported slide has no shape tree.');
    let id = Math.max(0, ...elements(doc, 'cNvPr').map((e) => Number(e.getAttribute('id')) || 0));
    for (const element of additions) {
      const shape = makeShape(doc, element, ++id, width, height);
      tree.insertBefore(
        shape,
        Array.from(tree.children).find((e) => e.localName === 'extLst') || null,
      );
    }
    zip.file(slide.sourcePath!, new XMLSerializer().serializeToString(doc), {
      createFolders: false,
    });
  }
  return new Blob([await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })], {
    type: MIME,
  });
}

function makeShape(doc: Document, item: SlideElement, id: number, width: number, height: number) {
  function node(
    parent: Element | null,
    ns: string,
    name: string,
    attrs: Record<string, string | number> = {},
  ) {
    const element = doc.createElementNS(ns, (ns === P ? 'p:' : 'a:') + name);
    for (const [key, value] of Object.entries(attrs)) element.setAttribute(key, String(value));
    parent?.appendChild(element);
    return element;
  }
  const shape = node(null, P, 'sp');
  const nv = node(shape, P, 'nvSpPr');
  node(nv, P, 'cNvPr', { id, name: `Noffice ${item.type} ${id}` });
  node(nv, P, 'cNvSpPr', item.type === 'text' ? { txBox: 1 } : {});
  node(nv, P, 'nvPr');
  const properties = node(shape, P, 'spPr');
  const transform = node(properties, A, 'xfrm');
  if (item.rotation) transform.setAttribute('rot', String(Math.round(item.rotation * 60000)));
  node(transform, A, 'off', {
    x: Math.round((item.x * width) / 960),
    y: Math.round((item.y * height) / 540),
  });
  node(transform, A, 'ext', {
    cx: Math.round((item.w * width) / 960),
    cy: Math.round((item.h * height) / 540),
  });
  node(
    node(properties, A, 'prstGeom', { prst: item.type === 'ellipse' ? 'ellipse' : 'rect' }),
    A,
    'avLst',
  );
  if (/^#[a-f\d]{6}$/i.test(item.fill)) {
    const color = node(node(properties, A, 'solidFill'), A, 'srgbClr', { val: item.fill.slice(1) });
    if ((item.fillOpacity ?? 1) < 1)
      node(color, A, 'alpha', { val: Math.round(item.fillOpacity! * 100000) });
  } else node(properties, A, 'noFill');
  if (item.outline) patchOutline(properties, item, width);
  else node(node(properties, A, 'ln'), A, 'noFill');
  const body = node(shape, P, 'txBody');
  node(body, A, 'bodyPr', { wrap: 'square', lIns: 0, rIns: 0, tIns: 0, bIns: 0 });
  node(body, A, 'lstStyle');
  for (const line of item.text.split('\n')) {
    const paragraph = node(body, A, 'p');
    node(paragraph, A, 'pPr', { algn: { left: 'l', center: 'ctr', right: 'r' }[item.align] });
    const run = node(paragraph, A, 'r');
    const format = node(run, A, 'rPr', {
      lang: 'en-US',
      sz: Math.round((item.fontSize * width) / 960 / 127),
      b: item.bold ? 1 : 0,
      i: item.italic ? 1 : 0,
      u: item.underline ? 'sng' : 'none',
    });
    node(node(format, A, 'solidFill'), A, 'srgbClr', {
      val: /^#[a-f\d]{6}$/i.test(item.color) ? item.color.slice(1) : '000000',
    });
    node(format, A, 'latin', { typeface: item.fontFamily || 'Arial' });
    node(run, A, 't').textContent = line;
  }
  return shape;
}
