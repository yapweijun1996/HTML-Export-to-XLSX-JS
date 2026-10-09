import http from 'node:http';
import path from 'node:path';
import { readFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const port = Number(process.env.PORT || 19483);
const types = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
    '.svg': 'image/svg+xml', '.json': 'application/json' };
http.createServer(async (request, response) => {
    try {
        const url = new URL(request.url, 'http://127.0.0.1:' + port);
        if (url.pathname === '/baseline.js') {
            const source = execFileSync('git', ['show', '4d93f11:export_to_xlsx.js'], { cwd: root });
            response.writeHead(200, { 'content-type': types['.js'] }); response.end(source); return;
        }
        const target = path.resolve(root, '.' + decodeURIComponent(url.pathname));
        if (!target.startsWith(root)) { response.writeHead(403); response.end(); return; }
        const content = await readFile(target);
        response.writeHead(200, { 'content-type': types[path.extname(target)] || 'application/octet-stream' });
        response.end(content);
    } catch { response.writeHead(404); response.end('Not found'); }
}).listen(port, '127.0.0.1', () => console.log('Fixture server: http://127.0.0.1:' + port));
