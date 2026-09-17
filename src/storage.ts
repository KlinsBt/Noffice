import type { OfficeFile } from './model';

const DB_NAME = 'noffice-workspace';
let connection: Promise<IDBDatabase> | undefined;
export function openDatabase(): Promise<IDBDatabase> {
  if (connection) return connection;
  connection = new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    let blocked = false;
    request.onupgradeneeded = () => {
      const db = request.result;
      db.createObjectStore('files', { keyPath: 'id' });
      db.createObjectStore('versions', { keyPath: 'key' }).createIndex('fileId', 'fileId');
    };
    request.onerror = () => {
      connection = undefined;
      reject(request.error);
    };
    request.onblocked = () => {
      blocked = true;
      connection = undefined;
      reject(new Error('Close other Noffice tabs and reload to update local storage.'));
    };
    request.onsuccess = () => {
      const db = request.result;
      if (blocked) {
        db.close();
        return;
      }
      db.onversionchange = () => {
        db.close();
        connection = undefined;
      };
      resolve(db);
    };
  });
  return connection;
}
export async function listFiles(): Promise<OfficeFile[]> {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const request = db.transaction('files').objectStore('files').getAll();
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}
export async function saveFile(file: OfficeFile, expectedRevision: number): Promise<number> {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(['files', 'versions'], 'readwrite');
    let failure: Error | undefined;
    const revision = expectedRevision + 1;
    tx.oncomplete = () => resolve(revision);
    tx.onabort = tx.onerror = () =>
      reject(
        failure || tx.error || new Error('Local save failed. Export a backup before leaving.'),
      );
    const files = tx.objectStore('files');
    const request = files.get(file.id);
    request.onsuccess = () => {
      const current = request.result as OfficeFile | undefined;
      if ((current?.revision ?? 0) !== expectedRevision || (!current && expectedRevision !== 0)) {
        failure = new Error(
          'This file changed in another tab. Export your work as a backup, then reload before continuing.',
        );
        tx.abort();
        return;
      }
      const next = { ...file, revision };
      files.put(next);
      const versions = tx.objectStore('versions');
      versions.put({
        key: `${file.id}:${revision}`,
        fileId: file.id,
        revision,
        at: file.updatedAt,
        name: file.name,
        content: file.content,
      });
      if (revision > 20) versions.delete(`${file.id}:${revision - 20}`);
    };
  });
}
export interface Version {
  key: string;
  fileId: string;
  revision: number;
  at: number;
  name: string;
  content: OfficeFile['content'];
}
export async function getVersions(id: string): Promise<Version[]> {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const request = db.transaction('versions').objectStore('versions').index('fileId').getAll(id);
    request.onsuccess = () =>
      resolve(request.result.sort((a: Version, b: Version) => b.revision - a.revision));
    request.onerror = () => reject(request.error);
  });
}
export async function closeDatabase() {
  const db = await connection;
  db?.close();
  connection = undefined;
}
