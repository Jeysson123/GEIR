import { test } from 'node:test';
import assert from 'node:assert/strict';
import { computePlan, longDate, shortDate } from '../../src/domain/admin/estimates';

const PHASES = [
    { actividad: 'Refinamiento y Estimaciones', dias: 1 },
    { actividad: 'Análisis y Diseño (GEIR)', dias: 2 },
    { actividad: 'Desarrollo y Pruebas Unitarias', desarrollo: true, minimo: 1 },
    { actividad: 'Evaluación Capacidad', dias: 0 },
    { actividad: 'Pruebas QA', dias: 6 },
    { actividad: 'Liberación y Pase a Producción', dias: 1 }
];

test('reproduce el inicio del plan del GEIR de ejemplo', () => {
    // Primera historia el miércoles 17/12/2025: 1 día de refinamiento, 2 de análisis, desarrollo desde el lunes 22.
    const { rows } = computePlan(PHASES, '2025-12-17', '2026-01-15');
    assert.deepEqual(rows[0], { activity: 'Refinamiento y Estimaciones', days: 1, start: '2025-12-17', end: '2025-12-17' });
    assert.deepEqual(rows[1], { activity: 'Análisis y Diseño (GEIR)', days: 2, start: '2025-12-18', end: '2025-12-19' });
    assert.equal(rows[2].start, '2025-12-22');
    assert.equal(rows[2].end, '2026-01-15');
    assert.equal(rows[3].days, 0);
});

test('salta fines de semana y feriados', () => {
    const { rows } = computePlan(PHASES, '2026-10-02', '2026-10-02', ['2026-10-05']);
    // Viernes 2: refinamiento; lunes 5 es feriado → análisis martes 6 y miércoles 7.
    assert.equal(rows[0].end, '2026-10-02');
    assert.deepEqual([rows[1].start, rows[1].end], ['2026-10-06', '2026-10-07']);
    assert.equal(rows[2].days, 1, 'el desarrollo nunca queda en 0 días');
});

test('formatos de fecha de los correos', () => {
    assert.equal(shortDate('2026-10-06'), '06/10/2026');
    assert.equal(longDate('2026-10-06'), 'Martes 06 de octubre del 2026');
});
