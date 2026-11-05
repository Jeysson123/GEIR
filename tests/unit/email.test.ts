import { test } from 'node:test';
import assert from 'node:assert/strict';
import { approvalHtml, buildEml, forumHtml, loadEmailTemplates } from '../../src/domain/admin/email';
import { configAsset } from '../../src/domain/admin/config';

test('crea un borrador .eml que Outlook abre como mensaje sin enviar', () => {
    const eml = buildEml({
        to: ['Usuario Uno (Empresa) <Usuario_Uno@example.com>'],
        cc: ['Área de Información <grupo@example.com>'],
        subject: 'CRQ12345 - Aprobación Diseño de Solución (Técnico)',
        html: '<p>Hola</p>',
        attachments: [{ name: 'CRQ12345 - Diseño Solución (Técnico).xlsx', data: Buffer.from('xlsx') }]
    }).toString('utf8');
    assert.match(eml, /^X-Unsent: 1\r\n/);
    assert.match(eml, /To: "Usuario Uno \(Empresa\)" <Usuario_Uno@example\.com>/);
    assert.match(eml, /Cc: =\?UTF-8\?B\?[^?]+\?= <grupo@example\.com>/, 'nombres con tildes codificados');
    assert.match(eml, /Subject: =\?UTF-8\?B\?/);
    assert.match(eml, /Content-Disposition: attachment;/);
    assert.ok(!/^From:/m.test(eml), 'sin remitente: Outlook usa la cuenta del usuario');
});

const templates = loadEmailTemplates((file) => configAsset(`correos/${file}`));

test('el cuerpo es el correo de Outlook original con los datos del cambio', () => {
    const html = approvalHtml(templates, {
        greeting: 'Buenas tardes',
        geir: 'GEIR_XX000',
        document: 'de las Pruebas Técnicas (Bitácora de Pruebas)',
        change: 'CRQ12345',
        system: 'Sistema',
        rows: [
            { requirement: '2026-REQ-0101', description: '100002: Nuevo reporte', type: 'Funcionalidad' },
            { requirement: '2026-REQ-0102', description: '100001: Separar <botones>', type: 'Corrección Defecto' }
        ],
        schedule: 'Fecha del cambio: 06/10/2026.',
        signature: { name: 'Usuario Dos', area: 'Área', role: 'Desarrollador', phone: '', motto: 'Lema', logoCid: 'logo@x' }
    });
    assert.match(html, /Generator content="Microsoft Word 15/, 'conserva el HTML de Outlook');
    assert.match(html, /aprobación de las <b>Pruebas Técnicas \(Bitácora de Pruebas\) <\/b>/);
    assert.match(html, /background:#0E2841/, 'encabezado de la tabla igual al original');
    assert.equal((html.match(/<tr style='height:15.75pt'>/g) ?? []).length, 3, 'encabezado + una fila por requerimiento');
    assert.match(html, /100001: Separar &lt;botones&gt;/);
    assert.match(html, /src="cid:logo@x"/);
    assert.ok(!/\{\{/.test(html), 'sin campos sin llenar');
    assert.ok(!/Celular/.test(html), 'sin teléfono no queda la línea vacía');
    assert.ok(!/REQ-0201|CRQ12345 que/.test(html.replace(/CRQ12345<\/span>/, '')), 'nada del ejemplo');
});

test('el FORO agrupa funcionalidades y lista los requerimientos de tres en tres', () => {
    const html = forumHtml(templates, {
        greeting: 'Saludos Usuario Uno',
        system: 'Sistema',
        change: 'CRQ12345',
        items: ['1', '2', '3', '4'].map((n) => ({ requirement: `2026-REQ-000${n}`, functional: `F${n}`, technical: `T${n}`, type: 'Funcionalidad' })),
        status: 'En proceso',
        date: 'Martes 06 de octubre del 2026',
        time: '8:00 PM – 10:30 PM',
        engineers: 'Usuario Dos',
        specialist: 'Usuario Uno',
        owner: 'Usuario Tres',
        certifiers: 'Usuario Cuatro',
        backend: ['api.jar'],
        frontend: ['bundle.js'],
        signature: { name: 'Usuario Dos', area: '', role: '', phone: '', motto: '', logoCid: 'logo@x' }
    });
    assert.match(html, /2026-REQ-0001, 2026-REQ-0002, 2026-REQ-0003 ?<o:p>/);
    assert.match(html, /FUNCIONALIDADES/);
    assert.match(html, /api\.jar/);
    assert.match(html, /bundle\.js/);
    assert.ok(!/\{\{/.test(html));
});
