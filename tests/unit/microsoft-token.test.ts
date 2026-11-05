import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as path from 'path';
import { MicrosoftToken } from '../../src/infrastructure/microsoft/microsoft-token';

test('el token de Microsoft vive en memoria y avisa cuando se pierde', async () => {
    const seen: (string | null)[] = [];
    const empty = new MicrosoftToken('https://contoso-my.sharepoint.com', null, (t) => seen.push(t));
    assert.equal(await empty.token(), null, 'sin token hay que iniciar sesión (sin llamar a Microsoft)');

    const kept = new MicrosoftToken('https://contoso-my.sharepoint.com', 'refresh-del-navegador', (t) => seen.push(t));
    kept.forget();
    assert.deepEqual(seen, [null], 'el navegador debe borrar su copia');
    assert.equal(await kept.token(), null);
    await kept.dispose();
});

test('no quedan archivos de sesión compartidos en el código del servidor', () => {
    const src = path.resolve(__dirname, '../../src');
    const files: string[] = [];
    const walk = (dir: string): void =>
        fs.readdirSync(dir, { withFileTypes: true }).forEach((e) =>
            e.isDirectory() ? walk(path.join(dir, e.name)) : e.name.endsWith('.ts') && files.push(path.join(dir, e.name))
        );
    walk(src);
    for (const file of files) {
        const text = fs.readFileSync(file, 'utf8');
        assert.ok(!/(microsoft|gitlab)\.(storage|token)\.json/.test(text), `${file} no debe usar sesiones guardadas en disco`);
    }
});
