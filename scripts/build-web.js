// Bundles the React 15 UI (web/src) into web/dist/app.js.
//
// Uses the native esbuild binary when it may run. Machines whose security policy blocks
// esbuild.exe (spawn EPERM) fall back to esbuild-wasm, which runs inside node.exe.
const fs = require('fs');
const path = require('path');

const watch = process.argv.includes('--watch');
const forceWasm = process.argv.includes('--wasm');
const root = path.resolve(__dirname, '..');
const outfile = path.join(root, 'web/dist/app.js');

const options = {
    entryPoints: [path.join(root, 'web/src/index.jsx')],
    outfile,
    bundle: true,
    minify: !watch,
    sourcemap: watch,
    target: 'es2017',
    jsx: 'transform',
    loader: { '.js': 'jsx' },
    define: { 'process.env.NODE_ENV': JSON.stringify(watch ? 'development' : 'production') },
    logLevel: 'info'
};

function isBlocked(error) {
    return /EPERM|EACCES|ENOENT|spawn|blocked|not compatible/i.test(`${error && error.code} ${error && error.message}`);
}

async function run(esbuild) {
    if (watch) {
        const ctx = await esbuild.context(options);
        await ctx.watch();
        return;
    }
    await esbuild.build(options);
}

async function main() {
    if (!forceWasm) {
        try {
            await run(require('esbuild'));
            return;
        } catch (error) {
            if (!isBlocked(error)) {
                throw error;
            }
            console.warn(`esbuild nativo no se pudo ejecutar (${error.code || error.message}); usando esbuild-wasm.`);
        }
    }
    try {
        await run(require('esbuild-wasm'));
    } catch (error) {
        if (!watch && fs.existsSync(outfile)) {
            console.warn(`No se pudo compilar la interfaz (${error.message}); se usa el web/dist/app.js existente.`);
            return;
        }
        throw error;
    }
}

main().catch((error) => {
    console.error(error.message || error);
    process.exit(1);
});
