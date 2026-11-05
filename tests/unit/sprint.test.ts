import { test } from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { buildSprintWorkbook } from '../../src/domain/admin/sprint-list';
import { configAsset } from '../../src/domain/admin/config';

test('la lista del sprint usa la plantilla original e incluye historias sin requerimiento', async () => {
    const template = configAsset('plantillas/sprint.xlsx');
    assert.ok(template, 'config/plantillas/sprint.xlsx');
    const data = await buildSprintWorkbook(template!, [
        { requirement: '2026-REQ-0101', id: '100002', title: 'Nuevo reporte', description: '100002: Nuevo reporte', type: 'Funcionalidad', assignedTo: 'Usuario Demo' },
        { requirement: '', id: '100006', title: 'Sin requerimiento', description: '100006: Sin requerimiento', type: 'Corrección Defecto', assignedTo: '' }
    ]);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(data as any);
    const general = wb.getWorksheet('Sheet1')!;
    const mail = wb.getWorksheet('Sheet2')!;
    assert.equal(general.getCell('B4').text, '2026-REQ-0101');
    assert.equal(general.getCell('C5').value, 100006);
    assert.equal(general.getCell('B5').text, '', 'sin requerimiento queda vacío');
    assert.equal(general.getCell('B6').text, '', 'no quedan filas del ejemplo');
    assert.equal(general.getCell('E17').text, '100006 - Sin requerimiento');
    assert.equal(mail.getCell('C5').text, '100002: Nuevo reporte');
    assert.equal(mail.getCell('D6').text, 'Corrección Defecto');
    assert.equal(mail.getCell('B7').text, '');
    // Formato de la plantilla: encabezado azul marino y bordes en las filas nuevas.
    assert.equal((mail.getCell('B4').fill as any)?.fgColor?.argb, 'FF0E2841');
    assert.ok(mail.getCell('C6').border?.left, 'las filas nuevas llevan los bordes del original');
});
