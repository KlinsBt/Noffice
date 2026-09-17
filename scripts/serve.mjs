import http from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import path from 'node:path';

// Test/preview helper only. The deployed application consists solely of build/ assets.
const root = path.resolve('build');
const types = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.txt': 'text/plain; charset=utf-8',
  '.woff2': 'font/woff2',
};
const server = http.createServer(async (request, response) => {
  try {
    if (!['GET', 'HEAD'].includes(request.method)) {
      response.writeHead(405).end();
      return;
    }
    const pathname = decodeURIComponent(new URL(request.url, 'http://localhost').pathname);
    let target = path.resolve(root, '.' + pathname);
    if (target !== root && !target.startsWith(root + path.sep)) {
      response.writeHead(403).end();
      return;
    }
    try {
      if ((await stat(target)).isDirectory()) target = path.join(target, 'index.html');
    } catch {
      if (!path.extname(pathname)) target = path.join(root, 'index.html');
    }
    const bytes = await readFile(target);
    response.writeHead(200, {
      'Content-Type': types[path.extname(target)] || 'application/octet-stream',
      'Cache-Control': 'no-cache',
      'X-Content-Type-Options': 'nosniff',
    });
    response.end(request.method === 'HEAD' ? undefined : bytes);
  } catch {
    response.writeHead(404).end('Not found');
  }
});
server.listen(Number(process.env.PORT || 4183), '127.0.0.1', () =>
  process.stdout.write(`Static Noffice preview at http://127.0.0.1:${process.env.PORT || 4183}\n`),
);
for (const signal of ['SIGINT', 'SIGTERM'])
  process.on(signal, () => server.close(() => process.exit(0)));
