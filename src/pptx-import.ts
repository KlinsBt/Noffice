import JSZip from 'jszip';
import { pptxStackNodes } from './pptx-stack';
import { groupTransform, projectGroupElement } from './pptx-group-transform';
import { presentationText } from './pptx-text';
import { presentationPaints, drawingChild } from './drawing-colors';
import { newSlide, textElement, uid, type SlideElement, type DeckContent } from './model';
function xml(text: string): XMLDocument {
  if (/<!DOCTYPE|<!ENTITY/i.test(text))
    throw new Error('XML entity declarations are not supported.');
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length)
    throw new Error('This Office file contains invalid XML.');
  return doc;
}
const all = (node: Document | Element, name: string) => [...node.getElementsByTagNameNS('*', name)];
function resolvePart(base: string, target: string) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const part of target.split('/')) {
    if (part === '..') parts.pop();
    else if (part !== '.') parts.push(part);
  }
  return parts.join('/');
}
export async function presentationRelations(zip: JSZip, part: string) {
  const path = part.replace(/([^/]+)$/, '_rels/$1.rels');
  const data = await zip.file(path)?.async('string');
  return new Map(
    data
      ? all(xml(data), 'Relationship')
          .filter((el) => el.getAttribute('TargetMode') !== 'External')
          .map((el) => [el.getAttribute('Id')!, resolvePart(part, el.getAttribute('Target')!)])
      : [],
  );
}

export async function readPresentation(
  data: ArrayBuffer,
  legacyText = false,
  legacyGroups = legacyText,
): Promise<DeckContent> {
  const zip = await JSZip.loadAsync(data),
    presentation = await zip.file('ppt/presentation.xml')?.async('string');
  if (!presentation) throw new Error('This file has no PowerPoint presentation.');
  const doc = xml(presentation),
    rels = await presentationRelations(zip, 'ppt/presentation.xml');
  const size = all(doc, 'sldSz')[0],
    width = Number(size?.getAttribute('cx') || 12192000),
    height = Number(size?.getAttribute('cy') || 6858000);
  if (
    !Number.isFinite(width / height) ||
    width <= 0 ||
    height <= 0 ||
    width / height < 0.1 ||
    width / height > 10
  )
    throw new Error('This presentation has unsupported slide dimensions.');
  const slides = [];
  for (const id of all(doc, 'sldId')) {
    const path = rels.get(
      id.getAttributeNS(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'id',
      ) ||
        id.getAttribute('r:id') ||
        '',
    );
    if (!path) throw new Error('A slide relationship is missing.');
    const text = await zip.file(path)?.async('string');
    if (!text) throw new Error('A slide is missing from this file.');
    const slideDoc = xml(text),
      slide = newSlide('blank'),
      slideRels = await presentationRelations(zip, path);
    const layoutPath = [...slideRels.values()].find((p) => p.includes('/slideLayouts/'));
    const layoutText = layoutPath && (await zip.file(layoutPath)?.async('string'));
    const layout = layoutText ? xml(layoutText) : null;
    const masterPath =
      layoutPath &&
      [...(await presentationRelations(zip, layoutPath)).values()].find((p) =>
        p.includes('/slideMasters/'),
      );
    const masterText = masterPath && (await zip.file(masterPath)?.async('string'));
    const master = masterText ? xml(masterText) : null;
    const themePath =
      masterPath &&
      [...(await presentationRelations(zip, masterPath)).values()].find((p) =>
        p.includes('/theme/'),
      );
    const themeText = themePath && (await zip.file(themePath)?.async('string'));
    const theme = themeText ? xml(themeText) : null;
    const paints = presentationPaints(theme, master, layout, slideDoc);
    const placeholder = (shape: Element, source: Document | null, byType = false) => {
      const ph = all(shape, 'ph')[0];
      if (!ph || !source) return undefined;
      const type = (p: Element) => {
        const t = p.getAttribute('type') || 'obj';
        return t === 'ctrTitle' ? 'title' : t;
      };
      const shapes = all(source, 'sp').filter((s) => all(s, 'ph').length);
      return byType
        ? shapes.find((s) => type(all(s, 'ph')[0]) === type(ph))
        : shapes.find(
            (s) => (all(s, 'ph')[0].getAttribute('idx') || '0') === (ph.getAttribute('idx') || '0'),
          );
    };
    const background = paints.background();
    slide.background = background.color === 'transparent' ? '#ffffff' : background.color;
    slide.backgroundOpacity = background.opacity;
    slide.sourcePath = path;
    slide.sourceLayoutPath = layoutPath || undefined;
    slide.sourceMasterPath = masterPath || undefined;
    slide.sourceThemePath = themePath || undefined;
    const tree = all(slideDoc, 'spTree')[0];
    const stack = tree ? pptxStackNodes(tree) : new Map<string, Element>();
    slide.stackOrder = [...stack.keys()];
    const shapeKeys = new Map<Element, string>();
    for (const [key, node] of stack) {
      shapeKeys.set(node, key);
      for (const child of all(node, '*')) shapeKeys.set(child, key);
    }
    for (const shape of all(slideDoc, '*').filter((node) =>
      ['sp', 'pic'].includes(node.localName),
    )) {
      const layoutShape = placeholder(shape, layout),
        masterShape = placeholder(layoutShape || shape, master, true);
      const transform =
          all(shape, 'xfrm')[0] ||
          (layoutShape && all(layoutShape, 'xfrm')[0]) ||
          (masterShape && all(masterShape, 'xfrm')[0]),
        offset = transform && all(transform, 'off')[0],
        extent = transform && all(transform, 'ext')[0];
      const sourceText = presentationText(
        shape,
        layoutShape,
        masterShape,
        master,
        theme,
        doc,
        paints,
        width,
      );
      const paragraphs = sourceText.text;
      const first = sourceText.base;
      const groupIds: string[] = [];
      for (
        let parent = shape.parentElement;
        parent?.localName === 'grpSp';
        parent = parent.parentElement
      ) {
        groupIds.unshift(all(parent, 'cNvPr')[0]?.getAttribute('id') || '');
        if (groupIds.length > 100) throw Error('This presentation has excessively nested groups.');
      }
      const ph = all(shape, 'ph')[0];
      const fill = paints.shape([shape, layoutShape, masterShape]);
      const element: SlideElement = textElement(paragraphs, {
        sourceShapeId: all(shape, 'cNvPr')[0]?.getAttribute('id') || uid(),
        sourceStackKey: shapeKeys.get(shape),
        sourceGroupIds: groupIds.length ? groupIds : undefined,
        sourcePlaceholder: ph
          ? {
              index: ph.getAttribute('idx') || '0',
              type: ph.getAttribute('type') || 'obj',
              layoutShapeId:
                (layoutShape && all(layoutShape, 'cNvPr')[0]?.getAttribute('id')) || undefined,
              masterShapeId:
                (masterShape && all(masterShape, 'cNvPr')[0]?.getAttribute('id')) || undefined,
              geometry: all(shape, 'xfrm')[0]
                ? 'slide'
                : layoutShape && all(layoutShape, 'xfrm')[0]
                  ? 'layout'
                  : masterShape && all(masterShape, 'xfrm')[0]
                    ? 'master'
                    : 'default',
            }
          : undefined,
        sourceText: paragraphs ? sourceText : undefined,
        rotation: (((Number(transform?.getAttribute('rot') || 0) / 60000) % 360) + 360) % 360,
        x: (Number(offset?.getAttribute('x') || 0) / width) * 960,
        y: (Number(offset?.getAttribute('y') || 0) / height) * 540,
        w: Math.max(10, (Number(extent?.getAttribute('cx') || width * 0.8) / width) * 960),
        h: Math.max(10, (Number(extent?.getAttribute('cy') || height * 0.2) / height) * 540),
        fontSize: Math.min(160, Math.max(8, first.fontSize)),
        bold: first.bold,
        italic: first.italic,
        underline: first.underline,
        fontFamily: first.fontFamily,
        color: first.color,
        align: first.align,
        fill: fill?.color || 'transparent',
        fillOpacity: fill?.opacity ?? 1,
        outline: paints.outline([shape, layoutShape, masterShape], width),
      });
      if (legacyText) {
        // Frozen pre-inheritance projection, used only to distinguish untouched legacy values.
        const run = all(shape, 'rPr')[0];
        Object.assign(element, {
          text: all(shape, 'p')
            .map((p) =>
              all(p, 't')
                .map((t) => t.textContent)
                .join(''),
            )
            .join('\n'),
          fontSize: Math.min(
            160,
            Math.max(8, (Number(run?.getAttribute('sz') || 2400) * 127 * 960) / width),
          ),
          bold: run?.getAttribute('b') === '1',
          italic: run?.getAttribute('i') === '1',
          underline: !!run?.getAttribute('u') && run?.getAttribute('u') !== 'none',
          fontFamily: run ? all(run, 'latin')[0]?.getAttribute('typeface') || undefined : undefined,
          color:
            paints.color(drawingChild(run, 'solidFill')?.firstElementChild)?.color || '#263d34',
          align: 'left',
          sourceText: undefined,
        });
      }
      const blip = all(shape, 'blip')[0];
      if (blip) {
        const target = slideRels.get(blip.getAttribute('r:embed') || '');
        const image = target && zip.file(target),
          extension = target?.split('.').pop()?.toLowerCase();
        if (image && ['png', 'jpg', 'jpeg', 'gif', 'webp'].includes(extension || '')) {
          element.type = 'image';
          element.flipH = ['1', 'true'].includes(transform?.getAttribute('flipH') || '');
          element.flipV = ['1', 'true'].includes(transform?.getAttribute('flipV') || '');
          element.src = `data:image/${extension === 'jpg' ? 'jpeg' : extension};base64,${await image.async('base64')}`;
        } else continue;
      } else if (!paragraphs) {
        element.type =
          all(shape, 'prstGeom')[0]?.getAttribute('prst') === 'ellipse' ? 'ellipse' : 'rect';
      }
      if (!legacyGroups) {
        const matrix = groupTransform(shape);
        if (matrix) {
          // Group-local units may be tiny compared with slide EMUs. Applying the old
          // minimum-size display clamp before scaling can enlarge a leaf thousands-fold.
          const localWidth = Number(extent?.getAttribute('cx') || width * 0.8);
          const localHeight = Number(extent?.getAttribute('cy') || height * 0.2);
          if (localWidth > 0) element.w = (localWidth / width) * 960;
          if (localHeight > 0) element.h = (localHeight / height) * 540;
          projectGroupElement(element, matrix, width, height);
        }
      }
      slide.elements.push(element);
    }
    const notesPath = [...slideRels.values()].find((p) =>
      /notesSlides\/notesSlide\d+\.xml/.test(p),
    );
    const notes = notesPath && (await zip.file(notesPath)?.async('string'));
    if (notes)
      slide.notes = all(xml(notes), 'sp')
        .filter((sp) => all(sp, 'ph')[0]?.getAttribute('type') === 'body')
        .flatMap((sp) =>
          all(sp, 'p').map((p) =>
            all(p, 't')
              .map((t) => t.textContent)
              .join(''),
          ),
        )
        .join('\n');
    slides.push(slide);
  }
  if (!slides.length) throw new Error('This presentation has no slides.');
  return {
    kind: 'powerpoint',
    textModelVersion: legacyText ? undefined : 1,
    groupModelVersion: legacyGroups ? undefined : 1,
    aspectRatio: width / height,
    slides,
  };
}
