import { z } from 'zod';
import { wordSectionIdSchema } from './word-section-identity';

export const isWordStoryPath = (path: string) =>
  !/[\\:?#\u0000-\u001f\u007f]/.test(path) &&
  path.split('/').every((part) => !!part && part !== '.' && part !== '..');

const storySectionId = wordSectionIdSchema;
export const wordStorySectionOptionsSchema = z.object({
  sectionId: storySectionId,
  differentFirstPage: z.boolean().optional(),
  headerDistance: z.number().int().min(0).max(31680).optional(),
  footerDistance: z.number().int().min(0).max(31680).optional(),
});

const referenceScope = {
  sectionId: storySectionId,
  kind: z.enum(['header', 'footer']),
  slot: z.enum(['default', 'first', 'even']),
};
export const wordStoryReferenceSchema = z.discriminatedUnion('linked', [
  z.object({ ...referenceScope, linked: z.literal(true) }),
  z.object({
    ...referenceScope,
    linked: z.literal(false),
    relationshipId: z.string().min(1).max(200),
  }),
]);

/** Header/footer contents are independent document stories. Their package
 * identities remain distinct from body paragraphs and section references. */
export const wordStoriesSchema = z
  .object({
    version: z.literal(1),
    evenAndOddHeaders: z.boolean(),
    templateVersion: z.union([z.literal(2), z.literal(3), z.literal(4)]).optional(),
    emptyTemplates: z
      .object({
        header: z.string().max(20000).nullable(),
        footer: z.string().max(20000).nullable(),
      })
      .optional(),
    sectionOptions: z.array(wordStorySectionOptionsSchema).max(10000).optional(),
    references: z.array(wordStoryReferenceSchema).max(60000).optional(),
    parts: z
      .array(
        z.object({
          path: z.string().min(1).max(500).refine(isWordStoryPath, 'Invalid Word story part path.'),
          kind: z.enum(['header', 'footer']),
          created: z.literal(true).optional(),
          copiedFrom: z.string().min(1).max(500).refine(isWordStoryPath).optional(),
          relationshipIds: z.array(z.string().min(1).max(200)).min(1).max(600),
          html: z.string().max(1000000),
        }),
      )
      .max(600),
  })
  .superRefine((stories, context) => {
    const sections = new Set<string>();
    for (const option of stories.sectionOptions || []) {
      if (sections.has(option.sectionId))
        context.addIssue({ code: 'custom', message: 'Duplicate story section options.' });
      sections.add(option.sectionId);
    }
    const paths = new Set<string>(),
      relationships = new Set<string>();
    let characters = 0;
    for (const part of stories.parts) {
      if (paths.has(part.path))
        context.addIssue({ code: 'custom', message: 'Duplicate Word story part.' });
      paths.add(part.path);
      for (const id of part.relationshipIds) {
        if (relationships.has(id))
          context.addIssue({ code: 'custom', message: 'Ambiguous Word story relationship.' });
        relationships.add(id);
      }
      characters += part.html.length;
    }
    if (characters > 10000000)
      context.addIssue({ code: 'custom', message: 'Word stories exceed the supported size.' });
    const slots = new Set<string>();
    for (const reference of stories.references || []) {
      const key = `${reference.sectionId}:${reference.kind}:${reference.slot}`;
      if (slots.has(key))
        context.addIssue({ code: 'custom', message: 'Duplicate story reference override.' });
      slots.add(key);
      if (
        !reference.linked &&
        !stories.parts.some(
          (p) => p.kind === reference.kind && p.relationshipIds.includes(reference.relationshipId),
        )
      )
        context.addIssue({ code: 'custom', message: 'Story reference has no matching part.' });
    }
    for (const part of stories.parts) {
      if (!part.copiedFrom && !part.created) continue;
      const source = stories.parts.find((p) => p.path === part.copiedFrom);
      const match = /^word\/noffice-(header|footer)-([a-f0-9]{32})\.xml$/.exec(part.path);
      if (
        (part.created
          ? !!part.copiedFrom
          : !source || source.copiedFrom || source.created || source.kind !== part.kind) ||
        !match ||
        match[1] !== part.kind ||
        part.relationshipIds.length !== 1 ||
        part.relationshipIds[0] !== `rIdNofficeStory${match[2]}`
      )
        context.addIssue({ code: 'custom', message: 'Invalid cloned story provenance.' });
    }
  });

export type WordStories = z.infer<typeof wordStoriesSchema>;
export type WordStoryReference = z.infer<typeof wordStoryReferenceSchema>;
