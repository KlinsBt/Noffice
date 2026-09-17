import { afterEach, describe, expect, it } from 'vitest';
import { newFile } from './model';
import { closeDatabase, getVersions, listFiles, saveFile } from './storage';
afterEach(async () => {
  await closeDatabase();
  await new Promise<void>((resolve, reject) => {
    const r = indexedDB.deleteDatabase('noffice-workspace');
    r.onsuccess = () => resolve();
    r.onerror = () => reject(r.error);
  });
});
describe('transactional local persistence', () => {
  it('persists files and content across database connections', async () => {
    const file = newFile('word');
    expect(await saveFile(file, 0)).toBe(1);
    await closeDatabase();
    const files = await listFiles();
    expect(files).toHaveLength(1);
    expect(files[0].content).toEqual(file.content);
    expect(files[0].revision).toBe(1);
  });
  it('rejects a stale tab without overwriting newer work', async () => {
    const file = newFile('excel');
    await saveFile(file, 0);
    await saveFile({ ...file, name: 'Newer work' }, 1);
    await expect(saveFile({ ...file, name: 'Stale work' }, 1)).rejects.toThrow('another tab');
    expect((await listFiles())[0].name).toBe('Newer work');
  });
  it('serializes concurrent revision checks so only one writer wins', async () => {
    const file = newFile('powerpoint');
    await saveFile(file, 0);
    const results = await Promise.allSettled([
      saveFile({ ...file, name: 'A' }, 1),
      saveFile({ ...file, name: 'B' }, 1),
    ]);
    expect(results.filter((r) => r.status === 'fulfilled')).toHaveLength(1);
    expect((await listFiles())[0].revision).toBe(2);
  });
  it('keeps the last 20 versions and can restore an earlier model', async () => {
    const file = newFile('word');
    for (let i = 0; i < 24; i++) await saveFile({ ...file, name: `Version ${i + 1}` }, i);
    const versions = await getVersions(file.id);
    expect(versions).toHaveLength(20);
    expect(versions[0].revision).toBe(24);
    expect(versions[19].revision).toBe(5);
    await saveFile({ ...file, name: versions[19].name, content: versions[19].content }, 24);
    expect((await listFiles())[0].name).toBe('Version 5');
  });
  it('keeps trash recoverable', async () => {
    const file = newFile('word');
    await saveFile(file, 0);
    await saveFile({ ...file, trashed: true }, 1);
    expect((await listFiles())[0].trashed).toBe(true);
    await saveFile({ ...file, trashed: false }, 2);
    expect((await listFiles())[0].trashed).toBe(false);
  });
});
