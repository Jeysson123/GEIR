/**
 * E-mail DRAFTS as .eml files (RFC 5322 + MIME). The "X-Unsent: 1" header makes Outlook
 * open them as a new message ready to review and send. GEIR never sends anything.
 */

export interface EmailAttachment {
    name: string;
    data: Buffer;
    contentType?: string;
}

export interface InlineImage {
    cid: string;
    name: string;
    data: Buffer;
    contentType: string;
}

export interface EmailDraft {
    to: string[];
    cc: string[];
    subject: string;
    html: string;
    inline?: InlineImage[];
    attachments?: EmailAttachment[];
}

const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

export function buildEml(draft: EmailDraft): Buffer {
    const mixed = boundary('mixed');
    const related = boundary('related');
    const lines: string[] = [
        'X-Unsent: 1',
        `To: ${draft.to.map(address).join(', ')}`,
        ...(draft.cc.length ? [`Cc: ${draft.cc.map(address).join(', ')}`] : []),
        `Subject: ${encodeWord(draft.subject)}`,
        `Date: ${new Date().toUTCString().replace('GMT', '+0000')}`,
        'MIME-Version: 1.0',
        `Content-Type: multipart/mixed; boundary="${mixed}"`,
        '',
        `--${mixed}`,
        `Content-Type: multipart/related; boundary="${related}"`,
        '',
        `--${related}`,
        'Content-Type: text/html; charset="utf-8"',
        'Content-Transfer-Encoding: base64',
        '',
        base64Lines(Buffer.from(draft.html, 'utf8'))
    ];
    for (const image of draft.inline ?? []) {
        lines.push(
            `--${related}`,
            `Content-Type: ${image.contentType}; name="${image.name}"`,
            `Content-ID: <${image.cid}>`,
            `Content-Disposition: inline; filename="${image.name}"`,
            'Content-Transfer-Encoding: base64',
            '',
            base64Lines(image.data)
        );
    }
    lines.push(`--${related}--`);
    for (const file of draft.attachments ?? []) {
        lines.push(
            `--${mixed}`,
            `Content-Type: ${file.contentType ?? XLSX}; name="${encodeWord(file.name)}"`,
            `Content-Disposition: attachment; filename="${encodeWord(file.name)}"; filename*=UTF-8''${encodeURIComponent(file.name)}`,
            'Content-Transfer-Encoding: base64',
            '',
            base64Lines(file.data)
        );
    }
    lines.push(`--${mixed}--`, '');
    return Buffer.from(lines.join('\r\n'), 'utf8');
}

/** "Usuario Uno (Empresa) <Usuario_Uno@example.com>" with the name encoded if needed. */
function address(value: string): string {
    const m = /^\s*"?([^"<]*?)"?\s*<([^>]+)>\s*$/.exec(value);
    if (!m) {
        return value.trim();
    }
    const name = m[1].trim();
    if (!name) {
        return `<${m[2]}>`;
    }
    return /^[\x20-\x7e]*$/.test(name) ? `"${name.replace(/"/g, "'")}" <${m[2]}>` : `${encodeWord(name)} <${m[2]}>`;
}

/** RFC 2047 encoded word for non-ASCII header text. */
function encodeWord(text: string): string {
    return /^[\x20-\x7e]*$/.test(text) ? text : `=?UTF-8?B?${Buffer.from(text, 'utf8').toString('base64')}?=`;
}

function base64Lines(data: Buffer): string {
    return (data.toString('base64').match(/.{1,76}/g) ?? []).join('\r\n');
}

function boundary(kind: string): string {
    return `----=_GEIR_${kind}_${Date.now().toString(36)}${Math.random().toString(36).slice(2, 10)}`;
}

// ------------------------------------------------------------------ HTML bodies
//
// The bodies are the Outlook e-mails in config/correos/*.html (taken from the real approval
// and FORO messages), so the table, fonts and signature look exactly like the ones sent by
// hand. GEIR only fills the {{placeholders}}: {{name}} is escaped text, {{{name}}} is HTML.

export interface RequirementRow {
    requirement: string;
    description: string; // same text as the sprint list (Sheet2), e.g. "100005: Módulo Demo"
    type: string;
}

export interface Signature {
    name: string;
    area: string;
    role: string;
    phone: string;
    motto: string;
    logoCid: string;
}

export interface EmailTemplates {
    approval: string;
    approvalFirstRow: string;
    approvalRow: string;
    forum: string;
    forumRequirements: string;
    forumGroup: string;
    forumDescription: string;
    forumTechnical: string;
    forumBackend: string;
    forumFrontend: string;
}

const TEMPLATE_FILES: Record<keyof EmailTemplates, string> = {
    approval: 'aprobacion.html',
    approvalFirstRow: 'aprobacion-fila-primera.html',
    approvalRow: 'aprobacion-fila.html',
    forum: 'foro.html',
    forumRequirements: 'foro-requerimientos.html',
    forumGroup: 'foro-grupo.html',
    forumDescription: 'foro-descripcion.html',
    forumTechnical: 'foro-tecnico.html',
    forumBackend: 'foro-backend.html',
    forumFrontend: 'foro-frontend.html'
};

/** Reads the templates with `read(fileName)`; a missing file is reported by name. */
export function loadEmailTemplates(read: (file: string) => Buffer | null): EmailTemplates {
    const templates = {} as EmailTemplates;
    for (const [key, file] of Object.entries(TEMPLATE_FILES) as [keyof EmailTemplates, string][]) {
        const data = read(file);
        if (!data) {
            throw new Error(`Falta la plantilla de correo ${file} en la carpeta de configuración.`);
        }
        templates[key] = data.toString('utf8');
    }
    return templates;
}

export function escapeHtml(text: string): string {
    return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Fills {{name}} (escaped) and {{{name}}} (raw HTML). A paragraph whose only data is an
 * empty value (e.g. no phone in the signature) is removed, as if it were never written.
 */
export function renderTemplate(template: string, values: Record<string, string>): string {
    let html = template.replace(/<p\b[^>]*>(?:(?!<\/p>)[\s\S])*?\{\{(\w+)\}\}(?:(?!<\/p>)[\s\S])*?<\/p>/g, (paragraph, key) =>
        key in values && !values[key].trim() && !/\{\{\{/.test(paragraph) ? '' : paragraph
    );
    html = html.replace(/\{\{\{(\w+)\}\}\}/g, (_, key) => values[key] ?? '');
    return html.replace(/\{\{(\w+)\}\}/g, (_, key) => escapeHtml(values[key] ?? ''));
}

function signatureValues(sig: Signature): Record<string, string> {
    return {
        firmaNombre: sig.name,
        firmaArea: sig.area,
        firmaCargo: sig.role,
        firmaTelefono: sig.phone,
        firmaLema: sig.motto,
        logo: sig.logoCid
    };
}

/** "del Diseño de Solución (Técnico)" → ["del", "Diseño de Solución (Técnico)"]. */
export function splitArticle(document: string): [string, string] {
    const m = /^\s*(del|de la|de las|de los|de)\s+(.+)$/i.exec(document);
    return m ? [m[1], m[2].trim()] : ['del', document.trim()];
}

/** Body of an approval request (Diseño de Solución / Pruebas Técnicas / UAT). */
export function approvalHtml(
    t: EmailTemplates,
    options: {
        greeting: string;
        geir: string;
        document: string; // "del Diseño de Solución (Técnico)"
        change: string;
        system: string;
        rows: RequirementRow[];
        schedule: string; // with the date already in
        signature: Signature;
    }
): string {
    const [article, document] = splitArticle(options.document);
    const rows = options.rows
        .map((r, i) =>
            renderTemplate(i === 0 ? t.approvalFirstRow : t.approvalRow, {
                requerimiento: r.requirement,
                descripcion: r.description,
                tipo: r.type
            })
        )
        .join('');
    return renderTemplate(t.approval, {
        saludo: options.greeting,
        geir: options.geir,
        articulo: article,
        documento: document,
        cambio: options.change,
        sistema: options.system,
        filas: rows,
        programacion: options.schedule,
        ...signatureValues(options.signature)
    });
}

export interface ForumItem {
    requirement: string;
    functional: string;
    technical: string;
    type: string;
    withoutRequirement?: boolean; // labelled by its Azure DevOps ID, left out of "No. Req."
}

/** Body of the FORO (production changes) e-mail. */
export function forumHtml(
    t: EmailTemplates,
    options: {
        greeting: string; // "Saludos Usuario Uno"
        system: string;
        change: string;
        items: ForumItem[];
        status: string;
        date: string; // "Martes 06 de octubre del 2026"
        time: string;
        engineers: string;
        specialist: string;
        specialistLabel?: string;
        owner: string;
        certifiers: string;
        backend: string[];
        frontend: string[];
        signature: Signature;
    }
): string {
    const requirements = options.items.filter((i) => !i.withoutRequirement).map((i) => i.requirement);
    const lines: string[] = [];
    for (let i = 0; i < requirements.length; i += 3) {
        lines.push(renderTemplate(t.forumRequirements, { linea: requirements.slice(i, i + 3).join(', ') }));
    }
    const groups = new Map<string, ForumItem[]>();
    for (const item of options.items) {
        const key = /correcci/i.test(item.type) ? 'CORRECCIONES DE DEFECTOS' : 'FUNCIONALIDADES';
        groups.set(key, [...(groups.get(key) ?? []), item]);
    }
    const description = [...groups.entries()]
        .map(
            ([title, items]) =>
                renderTemplate(t.forumGroup, { titulo: title }) +
                items
                    .map((i) => renderTemplate(t.forumDescription, { requerimiento: i.requirement, texto: i.functional || '(sin diseño funcional)' }))
                    .join('')
        )
        .join('');
    const technical = options.items
        .map((i) => renderTemplate(t.forumTechnical, { requerimiento: i.requirement, texto: i.technical || '(sin diseño técnico)' }))
        .join('');
    return renderTemplate(t.forum, {
        saludo: options.greeting.replace(/,\s*$/, ''),
        sistema: options.system,
        cambio: options.change,
        requerimientos: lines.join(''),
        estatus: options.status,
        descripcion: description,
        tecnico: technical,
        backend: options.backend.map((f) => renderTemplate(t.forumBackend, { archivo: f })).join(''),
        frontend: options.frontend.map((f) => renderTemplate(t.forumFrontend, { archivo: f })).join(''),
        fecha: options.date,
        hora: options.time,
        ingenieros: options.engineers,
        especialista: options.specialist,
        etiquetaEspecialista: options.specialistLabel ?? 'Especialista',
        dueno: options.owner,
        certifica: options.certifiers,
        ...signatureValues(options.signature)
    });
}

/** "Buenos días" before noon, "Buenas tardes" until 7 PM, then "Buenas noches". */
export function greeting(now = new Date()): string {
    const h = now.getHours();
    return h < 12 ? 'Buenos días' : h < 19 ? 'Buenas tardes' : 'Buenas noches';
}
