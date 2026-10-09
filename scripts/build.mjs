import { build } from 'esbuild';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const pkg = JSON.parse(await readFile(new URL('../package.json', import.meta.url)));
const check = process.argv.includes('--check');
const banner = '/*! HTML Export to XLSX JS v' + pkg.version + ' | MIT | Copyright (c) 2026 YAP WEI JUN */';
await mkdir(new URL('../dist/', import.meta.url), { recursive: true });
const manifest = { version: pkg.version, files: {} };
for (const minify of [false, true]) {
    const filename = minify ? 'export_to_xlsx.min.js' : 'export_to_xlsx.js';
    const result = await build({
        absWorkingDir: root, entryPoints: ['src/browser.js'], bundle: true,
        format: 'iife', platform: 'browser', target: 'es2022',
        define: { __HTML_XLSX_VERSION__: JSON.stringify(pkg.version) },
        minify, write: false, legalComments: 'none', banner: { js: banner }
    });
    const bytes = result.outputFiles[0].contents;
    const path = new URL('../dist/' + filename, import.meta.url);
    if (check) {
        const current = await readFile(path);
        if (!current.equals(Buffer.from(bytes))) throw new Error('Stale generated file: ' + filename);
    } else await writeFile(path, bytes);
    manifest.files[filename] = { bytes: bytes.length, sha256: createHash('sha256').update(bytes).digest('hex') };
}
const manifestBytes = JSON.stringify(manifest, null, 2) + '\n';
const manifestPath = new URL('../dist/manifest.json', import.meta.url);
if (check) {
    if (await readFile(manifestPath, 'utf8') !== manifestBytes) throw new Error('Stale dist manifest');
} else await writeFile(manifestPath, manifestBytes);
console.log(check ? 'Generated files match source.' : JSON.stringify(manifest, null, 2));
