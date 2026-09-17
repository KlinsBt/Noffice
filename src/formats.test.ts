import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { contentSchema, newFile, samples } from './model';
import { inspectZip, nativeBackup, parseCSV, sanitizeHTML, writeCSV } from './formats';
describe('untrusted content boundaries', () => {
  it('removes scripts, event handlers, remote tracking images and unsafe URLs', () => {
    const safe = sanitizeHTML(
      '<script>alert(1)</script><p onclick="alert(1)">Text</p><img src="https://tracker.example/x"><a href="javascript:alert(1)">bad</a><img src="data:image/png;base64,AAAA">',
    );
    expect(safe).not.toMatch(/script|onclick|https:\/\/tracker|javascript:/);
    expect(safe).toContain('Text');
    expect(safe).toContain('data:image/png');
  });
  it('removes CSS network fetches', () => {
    expect(
      sanitizeHTML('<p style="background-image:url(https://example.com/x)">A</p>'),
    ).not.toContain('url(');
  });
  it('rejects malformed content models and future kinds', () => {
    expect(contentSchema.safeParse({ kind: 'excel', sheets: [] }).success).toBe(false);
    expect(contentSchema.safeParse({ kind: 'other' }).success).toBe(false);
    expect(contentSchema.safeParse({ kind: 'powerpoint', slides: [] }).success).toBe(false);
  });
  it('validates all provided templates', () =>
    samples().forEach((file) => expect(contentSchema.safeParse(file.content).success).toBe(true)));
  it('rejects invalid ZIP inputs', () =>
    expect(() => inspectZip(new Uint8Array(60).buffer)).toThrow('ZIP'));
  it('validates the central directory and rejects oversized advertised expansion', async () => {
    const zip = new JSZip().file('[Content_Types].xml', '<Types/>');
    const bytes = await zip.generateAsync({ type: 'arraybuffer' });
    expect(() => inspectZip(bytes)).not.toThrow();
    const v = new DataView(bytes),
      end = bytes.byteLength - 22,
      start = v.getUint32(end + 16, true);
    v.setUint32(start + 24, 101 * 1024 * 1024, true);
    expect(() => inspectZip(bytes)).toThrow('100 MB');
  });
  it('includes original bytes in a native backup', () => {
    const file = newFile('word');
    file.original = { name: 'test.docx', data: new Uint8Array([0, 128, 255]).buffer };
    const backup = JSON.parse(nativeBackup(file));
    expect(backup.version).toBe(1);
    expect([...atob(backup.original.base64)].map((c) => c.charCodeAt(0))).toEqual([0, 128, 255]);
    expect(backup.content).toEqual(file.content);
  });
});
describe('CSV interoperability', () => {
  it('handles quoted commas, embedded newlines, escaped quotes, BOM, and CRLF', () =>
    expect(parseCSV('\uFEFFA,B\r\n"hello, world","line\nnext"\r\n"say ""hi""",last')).toEqual([
      ['A', 'B'],
      ['hello, world', 'line\nnext'],
      ['say "hi"', 'last'],
    ]));
  it('preserves trailing empty fields', () =>
    expect(parseCSV('a,b,\n1,2,')).toEqual([
      ['a', 'b', ''],
      ['1', '2', ''],
    ]));
  it('rejects unterminated quoted fields', () =>
    expect(() => parseCSV('"hello')).toThrow('unclosed'));
  it('exports calculated values and protects text formula injection', () => {
    const s = {
      id: '1',
      name: 'One',
      rows: 100,
      cols: 26,
      cells: {
        A1: { value: '2' },
        B1: { value: '=A1*3' },
        A2: { value: '\'=HYPERLINK("https://example.com")' },
      },
    };
    const csv = writeCSV(s, [s]);
    expect(csv).toContain('2,6');
    expect(parseCSV(csv)[1][0]).toBe('\'=HYPERLINK("https://example.com")');
  });
});
