import { test } from 'node:test';
import assert from 'node:assert/strict';
import { descriptionLine, parseTaskLine, taskLabel } from '../../src/domain/document';

test('lee las líneas de tarea en todos los formatos del GEIR', () => {
    assert.deepEqual(parseTaskLine('7 - : 100007 : Módulo Demo: Auditoria'), {
        number: 7,
        requirement: '',
        id: '100007',
        title: 'Módulo Demo: Auditoria'
    });
    assert.deepEqual(parseTaskLine(': 100007: Módulo'), { number: null, requirement: '', id: '100007', title: 'Módulo' });
    assert.deepEqual(parseTaskLine('1 - 2026-REQ-0103: 100004 - Pantalla de consulta'), {
        number: 1,
        requirement: '2026-REQ-0103',
        id: '100004',
        title: 'Pantalla de consulta'
    });
    assert.equal(parseTaskLine('2026-REQ-0201-25: 100005 - Monitor')?.requirement, '2026-REQ-0201-25');
});

test('no confunde beneficios ni notas con tareas', () => {
    assert.equal(parseTaskLine('15 - Agregar filtros en los modulos'), null);
    assert.equal(parseTaskLine('Más detalles en Azure DevOps ID 100002'), null);
});

test('escribe el requerimiento delante solo cuando está relacionado', () => {
    assert.equal(descriptionLine(7, { id: '100007', title: 'Auditoria' }), '7 - : 100007 : Auditoria');
    assert.equal(
        descriptionLine(7, { id: '100007', title: 'Auditoria', requirement: '2026-REQ-0101' }),
        '7 - 2026-REQ-0101: 100007 - Auditoria'
    );
    assert.equal(taskLabel({ id: '100007', title: 'Auditoria', requirement: '2026-REQ-0101' }), '2026-REQ-0101: 100007 - Auditoria');
});
