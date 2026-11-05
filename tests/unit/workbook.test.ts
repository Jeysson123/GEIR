import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import { GeirDocument, emptyTask } from '../../src/domain/document';
import { XlsxPackage } from '../../src/domain/xlsx/package';

const TEMPLATE = fs.readFileSync(path.resolve(__dirname, '../../templates/N - GEIR.xlsx'));

async function withTask(): Promise<GeirDocument> {
    const doc = await GeirDocument.load(TEMPLATE);
    await doc.upsertTask({ ...emptyTask('100007'), title: 'Auditoria de accesos', benefit: 'Llevar control' });
    return GeirDocument.load(await doc.save());
}

test('relaciona el requerimiento en las cuatro hojas y llena el ID del Pedido', async () => {
    const doc = await withTask();
    await doc.setRequestIds(['2026-REQ-0101', '2026-REQ-0102']);
    assert.equal(await doc.relateRequirement('100007', '2026-REQ-0101'), true);
    assert.equal(await doc.relateRequirement('999999', '2026-REQ-0102'), false);
    const again = await GeirDocument.load(await doc.save());
    const [task] = await again.listTasks();
    assert.equal(task.requirement, '2026-REQ-0101');
    const read = await again.readTask('100007');
    assert.equal(read?.requirement, '2026-REQ-0101');
    assert.equal(read?.title, 'Auditoria de accesos');
    const sheet = await again.pkg.sheet('Descripción Requerimiento');
    assert.equal(await sheet.displayText('C7'), '2026-REQ-0101, 2026-REQ-0102');
});

test('llena el Plan de Trabajo y pide recalcular el total', async () => {
    const doc = await withTask();
    const missing = await doc.setWorkPlan([
        { activity: 'Refinamiento y Estimaciones', days: 1, start: '2025-12-17', end: '2025-12-17' },
        { activity: 'Evaluación Seguridad (Pre Prod)*', days: 0, start: '', end: '' },
        { activity: 'Actividad que no existe', days: 1, start: '2025-12-17', end: '2025-12-17' }
    ]);
    assert.deepEqual(missing, ['Actividad que no existe']);
    const zip = await JSZip.loadAsync(await doc.save());
    assert.match(await zip.file('xl/workbook.xml')!.async('string'), /fullCalcOnLoad="1"/);
});

test('deriva una sola hoja sin romper el paquete', async () => {
    const pkg = await XlsxPackage.load(TEMPLATE);
    await pkg.keepOnlySheet('Diseño Solución');
    const data = await pkg.save();
    const again = await XlsxPackage.load(data);
    assert.deepEqual(await again.sheetNames(), ['Diseño Solución']);
    const zip = await JSZip.loadAsync(data);
    assert.ok(zip.file('_rels/.rels'), 'conserva la relación raíz del paquete');
    // Every relationship points to a part that exists.
    for (const rels of Object.keys(zip.files).filter((n) => n.endsWith('.rels'))) {
        const owner = rels === '_rels/.rels' ? '' : path.posix.join(path.posix.dirname(path.posix.dirname(rels)), path.posix.basename(rels).replace(/\.rels$/, ''));
        const xml = await zip.file(rels)!.async('string');
        for (const m of xml.matchAll(/<Relationship [^>]*Target="([^"]+)"[^>]*>/g)) {
            if (/TargetMode="External"/.test(m[0])) {
                continue;
            }
            const target = m[1].startsWith('/') ? m[1].slice(1) : path.posix.normalize(path.posix.join(path.posix.dirname(owner), m[1]));
            assert.ok(zip.file(target), `${rels} apunta a ${target}, que no existe`);
        }
    }
});
