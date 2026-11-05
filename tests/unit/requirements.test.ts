import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseRequirementPairs } from '../../src/domain/admin/requirements';

test('acepta líneas simples, con título, filas del Excel y el orden inverso', () => {
    const { pairs, errors } = parseRequirementPairs(
        [
            '2026-REQ-0101 100002',
            '2026-REQ-0102: 100001 - Botones separados en el formulario',
            '3\t2026-REQ-0103\t100004\tPantalla de consulta\t100004 - Pantalla\tCorrección Defecto\tUsuario Demo',
            '100003 2026-REQ-0104',
            'Requerimiento\tID Azure DevOps\tDescripción'
        ].join('\n')
    );
    assert.deepEqual(errors, []);
    assert.deepEqual(
        pairs.map((p) => [p.requirement, p.storyId, p.title]),
        [
            ['2026-REQ-0101', '100002', ''],
            ['2026-REQ-0102', '100001', 'Botones separados en el formulario'],
            ['2026-REQ-0103', '100004', 'Pantalla de consulta'],
            ['2026-REQ-0104', '100003', '']
        ]
    );
});

test('acepta un arreglo JSON', () => {
    const { pairs } = parseRequirementPairs('[["2026-REQ-0101","100002"],{"requerimiento":"2026-REQ-0102","id":"100001"}]');
    assert.deepEqual(
        pairs.map((p) => p.storyId),
        ['100002', '100001']
    );
});

test('la relación es la indicada: avisa duplicados en vez de adivinar', () => {
    const { pairs, errors } = parseRequirementPairs('2026-REQ-0101 100002\n2026-REQ-0101 100001\n2026-REQ-0105 100002');
    assert.equal(pairs.length, 1);
    assert.equal(errors.length, 2);
});
