/** An export-only snapshot of settled semantic column boundaries. No geometry
 * enters document HTML, history or storage. The full content binding prevents
 * an asynchronous export from using another revision's layout. */
export interface WordExportLayout {
  content: string;
  paragraphs: { text: string; boundaries: number[] }[];
}
