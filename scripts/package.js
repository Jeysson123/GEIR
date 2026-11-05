// Builds the deployable bundle: release/geir-<version>/ and release/geir-<version>.zip.
// Run after `npm run build`. Same bundle for Windows and Linux: compiled JavaScript (no
// ts-node), the built interface, config/, templates/ and its own production node_modules.
// The target needs Node.js 18+ and a browser for the background sign-ins: Edge on Windows,
// Playwright's Chromium on Linux (see docs/ARCHITECTURE.md).
//
//   node scripts/package.js              full bundle (installs production dependencies)
//   node scripts/package.js --no-install  skips node_modules (faster, for checks)

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const JSZip = require('jszip');

const root = path.resolve(__dirname, '..');
const version = /VERSION = '([^']+)'/.exec(fs.readFileSync(path.join(root, 'src/server/version.ts'), 'utf8'))[1];
const name = `geir-${version}`;
const releaseDir = path.join(root, 'release');
const out = path.join(releaseDir, name);
const install = !process.argv.includes('--no-install');

function copy(from, to) {
    const src = path.join(root, from);
    if (!fs.existsSync(src)) {
        throw new Error(`Falta ${from}. ¿Ejecutaste "npm run build"?`);
    }
    fs.cpSync(src, path.join(out, to || from), { recursive: true });
}

console.log(`Empaquetando ${name}…`);
// Only the current version: older bundles are removed so nothing stale is left behind.
fs.rmSync(releaseDir, { recursive: true, force: true });
fs.mkdirSync(out, { recursive: true });

copy('build/src', 'src'); // compiled JavaScript, same folder layout as the sources
copy('web/index.html');
copy('web/styles.css');
copy('web/dist');
copy('config');
copy('templates');
copy('docs');
copy('README.md');
copy('package-lock.json');
copy('deploy');

const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
pkg.version = version.replace(/^(\d+)\.(\d+)\.(\d+).*$/, '$1.$2.$3');
pkg.scripts = {
    start: 'node src/server/server.js',
    demo: 'node src/server/server.js --mock',
    'browser:install': 'playwright install --with-deps chromium'
};
fs.writeFileSync(path.join(out, 'package.json'), JSON.stringify(pkg, null, 2) + '\n');

fs.writeFileSync(
    path.join(out, 'iniciar-geir.cmd'),
    [
        '@echo off',
        'rem GEIR: inicia el servidor en http://localhost:4300',
        'cd /d "%~dp0"',
        'node src\\app\\server.js',
        'pause',
        ''
    ].join('\r\n')
);
fs.writeFileSync(
    path.join(out, 'iniciar-geir.sh'),
    ['#!/usr/bin/env sh', '# GEIR (Linux): usa .env si existe', 'cd "$(dirname "$0")"', 'exec node src/server/server.js', ''].join('\n'),
    { mode: 0o755 }
);
fs.copyFileSync(path.join(root, '.env.example'), path.join(out, '.env.example'));

if (install) {
    console.log('Instalando dependencias de producción…');
    execSync('npm ci --omit=dev --ignore-scripts --no-audit --no-fund', {
        cwd: out,
        stdio: 'inherit',
        env: Object.assign({}, process.env, { PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD: '1' })
    });
}

/** Removes empty folders (npm leaves empty @scope folders of skipped dev packages). */
function removeEmptyDirs(dir) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        if (entry.isDirectory()) {
            removeEmptyDirs(path.join(dir, entry.name));
        }
    }
    if (dir !== out && fs.readdirSync(dir).length === 0) {
        fs.rmdirSync(dir);
    }
}
removeEmptyDirs(out);

async function zipFolder() {
    const zip = new JSZip();
    const add = (dir, prefix) => {
        for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
            const full = path.join(dir, entry.name);
            const inZip = `${prefix}/${entry.name}`;
            if (entry.isDirectory()) {
                add(full, inZip);
            } else {
                zip.file(inZip, fs.readFileSync(full));
            }
        }
    };
    add(out, name);
    const data = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    const file = path.join(releaseDir, `${name}.zip`);
    fs.writeFileSync(file, data);
    console.log(`Listo: ${path.relative(root, file)} (${(data.length / 1024 / 1024).toFixed(1)} MB)`);
}

zipFolder().catch((error) => {
    console.error(error);
    process.exit(1);
});
