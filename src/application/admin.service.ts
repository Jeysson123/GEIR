import * as path from 'path';
import { WorkItemSummary } from '../infrastructure/azure-devops/azure-devops.client';
import { DocumentLocation, StoredDocument, describeLocation, writeLocal } from '../infrastructure/microsoft/onedrive-files.client';
import { GeirDocument } from '../domain/document';
import { XlsxPackage } from '../domain/xlsx/package';
import {
    AdminConfig,
    configAsset,
    fill,
    loadAdminConfig,
    orTemporal,
    safeFileName,
    workTypeLabel
} from '../domain/admin/config';
import { parseRequirementPairs } from '../domain/admin/requirements';
import { computePlan, longDate, pickDate, shortDate } from '../domain/admin/estimates';
import { buildSprintWorkbook, SprintStory } from '../domain/admin/sprint-list';
import {
    ForumItem,
    RequirementRow,
    Signature,
    approvalHtml,
    buildEml,
    forumHtml,
    greeting,
    loadEmailTemplates
} from '../domain/admin/email';
import { mockSummary } from './mock-data';
import {
    DocumentInfo,
    MOCK,
    UserError,
    azure,
    editOnline,
    loadDocument,
    loadOrCreate,
    locate,
    saveOutlookDraft,
    withMicrosoft,
    writeDocument
} from './geir.service';
import { Session } from './session';

export interface AdminInput {
    location: string;
    cambio?: string; // change name, e.g. CRQ12345
    geir?: string; // GEIR name, e.g. GEIR_XX000-20260101-000000
    sprint?: string; // e.g. 24-2025
    requerimientos: string; // pairs REQ ↔ user story (see parseRequirementPairs)
    fechaPase?: string; // YYYY-MM-DD; default: release date of the work plan
}

export interface RequirementMatch {
    requirement: string;
    storyId: string;
    title: string;
    type: string;
    assignedTo: string;
    tasks: { id: string; number: number | null; title: string; how: 'misma' | 'hija' }[];
}

export interface AdminResult {
    names: { cambio: string; geir: string; sprint: string };
    matches: RequirementMatch[];
    warnings: string[];
    plan: {
        rows: { activity: string; days: number; start: string; end: string }[];
        totalDays: number;
        first: { id: string; date: string } | null;
        last: { id: string; date: string } | null;
        release: string;
    };
    files: { name: string; kind: 'geir' | 'derivado' | 'sprint' | 'correo'; size: number }[];
    /** E-mails saved in the Drafts folder of the user's Outlook (never sent). */
    drafts: { subject: string; to: string[]; saved: boolean; webLink: string; file: string }[];
    document: DocumentInfo;
}

/**
 * Admin options: relates requirements to the user stories of the GEIR, fills the
 * request IDs and the work plan, derives the per-approval workbooks and the sprint list,
 * and saves the e-mails as DRAFTS in the user's own Outlook mailbox. Nothing is ever sent.
 */
export async function adminGenerate(session: Session, input: AdminInput): Promise<AdminResult> {
    let config: AdminConfig;
    try {
        config = loadAdminConfig();
    } catch (error: any) {
        throw new UserError(error.message, 500);
    }
    const names = {
        cambio: orTemporal(config, input.cambio),
        geir: orTemporal(config, input.geir),
        sprint: orTemporal(config, input.sprint)
    };
    const { pairs, errors } = parseRequirementPairs(input.requerimientos ?? '');
    if (!pairs.length) {
        throw new UserError(
            errors[0] ?? 'Escribe los requerimientos y sus historias, uno por línea (ej. 2026-REQ-0101 100002).'
        );
    }
    const warnings = [...errors];
    const location = locate(input.location);

    // 1. What the GEIR has now, and what Azure says about each story and task.
    const { stored: original, doc: current } = await loadOrCreate(session, location);
    const tasks = await current.listTasks();
    const summaries = await summariesFor(session, [...pairs.map((p) => p.storyId), ...tasks.map((t) => t.id)], warnings);

    // 2. Requirement ↔ story ↔ GEIR task. Only exact relations: the task IS the story, or
    //    the task's parent in Azure DevOps is the story. Nothing is guessed.
    const matches: RequirementMatch[] = pairs.map((p) => {
        const story = summaries.get(p.storyId);
        return {
            requirement: p.requirement,
            storyId: p.storyId,
            title: story?.title || p.title || tasks.find((t) => t.id === p.storyId)?.title || '',
            type: workTypeLabel(config, story?.workItemType ?? ''),
            assignedTo: story?.assignedTo ?? '',
            tasks: []
        };
    });
    for (const task of tasks) {
        const parent = summaries.get(task.id)?.parentId ?? '';
        const direct = matches.find((m) => m.storyId === task.id);
        const byParent = direct ? null : matches.find((m) => parent && m.storyId === parent);
        const match = direct ?? byParent;
        if (match) {
            match.tasks.push({ id: task.id, number: task.number, title: task.title, how: direct ? 'misma' : 'hija' });
        } else {
            warnings.push(`La tarea ${task.id} del GEIR no pertenece a ninguna de las historias indicadas: se incluye sin requerimiento.`);
        }
    }
    for (const m of matches.filter((x) => !x.tasks.length)) {
        warnings.push(`${m.requirement}: la historia ${m.storyId} no tiene tareas en el GEIR (se incluye en los correos y en el sprint).`);
    }

    // 3. Work plan: from the first user story of the GEIR to the last one added.
    const ordered = matches
        .flatMap((m) => m.tasks.map((t) => ({ match: m, task: t })))
        .sort((a, b) => (a.task.number ?? Infinity) - (b.task.number ?? Infinity));
    const firstStory = ordered.length ? summaries.get(ordered[0].match.storyId) : summaries.get(matches[0].storyId);
    const lastStory = ordered.length ? summaries.get(ordered[ordered.length - 1].match.storyId) : summaries.get(matches[matches.length - 1].storyId);
    const firstDate = pickDate(firstStory, config.estimaciones.fechaPrimeraHistoria);
    const lastDate = pickDate(lastStory, config.estimaciones.fechaUltimaHistoria);
    if (!firstDate) {
        warnings.push('Azure DevOps no dio fecha para la primera historia: el plan empieza hoy.');
    }
    const plan = computePlan(config.estimaciones.fases, firstDate, lastDate, config.estimaciones.feriados);
    const release = /^\d{4}-\d{2}-\d{2}$/.test(input.fechaPase ?? '') ? input.fechaPase! : plan.release;

    // Tasks of the GEIR that belong to none of the stories: they stay, without requirement.
    const related = new Set(matches.flatMap((m) => m.tasks.map((t) => t.id)));
    const unmatched = tasks.filter((t) => !related.has(t.id));

    // 4. The GEIR itself: request IDs, requirement in front of each task, work plan, sprint.
    const update = async (doc: GeirDocument) => {
        await doc.setRequestIds(pairs.map((p) => p.requirement));
        for (const m of matches) {
            for (const t of m.tasks) {
                await doc.relateRequirement(t.id, m.requirement);
            }
        }
        for (const t of unmatched) {
            await doc.relateRequirement(t.id, null); // "ID - título", without the empty ": ID :"
        }
        await doc.tidyDescription();
        const missing = await doc.setWorkPlan(plan.rows);
        missing.forEach((a) => warnings.push(`El Plan de Trabajo del GEIR no tiene la actividad "${a}".`));
        if (input.sprint?.trim()) {
            await doc.setSprint(names.sprint.replace(/-/g, ' '));
        }
    };
    let geirData: Buffer;
    let stored: StoredDocument;
    if (location.kind === 'local') {
        await update(current);
        geirData = await current.save();
        stored = await writeDocument(session, location, geirData);
    } else {
        stored = await editOnline(session, location, async (data) => {
            const doc = await loadDocument(data ?? original.data!);
            await update(doc);
            geirData = await doc.save();
            return geirData;
        });
    }
    const doc = await loadDocument(geirData!);

    // 5. Files next to the GEIR.
    session.adminFiles.clear();
    const files: AdminResult['files'] = [];
    const values = { cambio: names.cambio, geir: names.geir, sprint: names.sprint, sistema: config.sistema };
    const put = async (name: string, data: Buffer, kind: AdminResult['files'][number]['kind']) => {
        const fileName = safeFileName(name);
        await writeSibling(session, location, fileName, data);
        session.adminFiles.set(fileName, data);
        files.push({ name: fileName, kind, size: data.length });
        console.log(`[GEIR] Generado ${fileName}`);
    };

    const geirCopy = fill(config.archivos.geir, values);
    if (safeFileName(geirCopy).toLowerCase() !== stored.name.toLowerCase()) {
        await put(geirCopy, geirData!, 'geir');
    }
    const derived = new Map<string, { name: string; data: Buffer }>();
    for (const d of config.archivos.derivados) {
        const pkg = await XlsxPackage.load(geirData!);
        await pkg.keepOnlySheet(d.hoja);
        const name = safeFileName(fill(d.archivo, values));
        const data = await pkg.save();
        derived.set(d.clave, { name, data });
        await put(name, data, 'derivado');
    }

    // The e-mail table is the sprint list's Sheet2: the same rows, the same text.
    const describe = (id: string, title: string) =>
        title ? fill(config.archivos.formatoDescripcion || '{id}: {titulo}', { id, titulo: title }) : id;
    // Every story of the document is listed, with or without requirement: first the stories
    // given with their requirement, then the rest of the GEIR in document order.
    const stories: SprintStory[] = matches.map((m) => ({
        requirement: m.requirement,
        id: m.storyId,
        title: m.title,
        description: describe(m.storyId, m.title),
        type: m.type,
        assignedTo: m.assignedTo
    }));
    for (const t of [...tasks].sort((a, b) => (a.number ?? Infinity) - (b.number ?? Infinity))) {
        if (stories.some((s) => s.id === t.id)) {
            continue;
        }
        const summary = summaries.get(t.id);
        const owner = matches.find((m) => m.tasks.some((x) => x.id === t.id));
        const title = summary?.title || t.title;
        stories.push({
            requirement: owner?.requirement ?? '',
            id: t.id,
            title,
            description: describe(t.id, title),
            type: workTypeLabel(config, summary?.workItemType ?? ''),
            assignedTo: summary?.assignedTo ?? ''
        });
    }
    const sprintTemplate = configAsset(config.archivos.plantillaSprint || 'plantillas/sprint.xlsx');
    if (!sprintTemplate) {
        throw new UserError('Falta la plantilla del sprint (config/plantillas/sprint.xlsx).', 500);
    }
    await put(fill(config.archivos.sprint, values), await buildSprintWorkbook(sprintTemplate, stories), 'sprint');

    // 6. E-mail drafts in the user's Outlook (Drafts folder): approvals with their workbook
    //    attached, and the FORO. If the mailbox cannot be reached, the draft is left as an
    //    .eml file next to the GEIR so nothing is lost.
    const drafts: AdminResult['drafts'] = [];
    let mailboxError = '';
    const draft = async (fileName: string, subject: string, to: string[], mime: Buffer) => {
        if (!mailboxError) {
            try {
                const saved = await saveOutlookDraft(session, mime);
                drafts.push({ subject, to, saved: true, webLink: saved.webLink, file: '' });
                console.log(`[GEIR] Borrador en Outlook: ${subject}`);
                return;
            } catch (error: any) {
                const reason = String(error?.message ?? error);
                warnings.push(`No se pudo guardar "${subject}" en tus borradores de Outlook (${reason}); quedó como archivo .eml en la carpeta del GEIR.`);
                if (!/^el correo pesa/.test(reason)) {
                    mailboxError = reason; // same problem for the rest: do not retry each one
                }
            }
        } else {
            warnings.push(`"${subject}" quedó como archivo .eml en la carpeta del GEIR (Outlook: ${mailboxError}).`);
        }
        await put(fileName, mime, 'correo');
        drafts.push({ subject, to, saved: false, webLink: '', file: safeFileName(fileName) });
    };
    const user = (session.displayName || session.credentials.username).replace(/\s*\([^)]*\)\s*$/, '');
    const people = { ...values, usuario: user, fecha: shortDate(release), fechaLarga: longDate(release) };
    const folder = config.correos.plantillas || 'correos';
    let templates;
    try {
        templates = loadEmailTemplates((file) => configAsset(`${folder}/${file}`));
    } catch (error: any) {
        throw new UserError(error.message, 500);
    }
    const firma = config.correos.firma;
    const logo = configAsset(firma.logo);
    const signature: Signature = {
        name: fill(firma.nombre, people),
        area: firma.area ?? '',
        role: firma.cargo ?? '',
        phone: firma.telefono ?? '',
        motto: firma.lema ?? '',
        logoCid: 'image001.png@geir.local'
    };
    const inline = logo ? [{ cid: signature.logoCid, name: 'image001.png', data: logo, contentType: 'image/png' }] : [];
    const rows: RequirementRow[] = stories.map((s) => ({ requirement: s.requirement, description: s.description, type: s.type }));
    const recipients = (list: string[]) => list.map((key) => config.personas[key] ?? key).filter(Boolean);
    const hello = greeting();

    for (const a of config.correos.aprobaciones) {
        const attachment = derived.get(a.adjunto);
        const html = approvalHtml(templates, {
            greeting: hello,
            geir: names.geir,
            document: a.documento,
            change: names.cambio,
            system: config.sistema,
            rows,
            schedule: fill(config.correos.programacion, people),
            signature
        });
        const subject = fill(a.asunto, people);
        const to = recipients(a.para);
        const eml = buildEml({
            to,
            cc: recipients(a.cc),
            subject,
            html,
            inline,
            attachments: attachment ? [{ name: attachment.name, data: attachment.data }] : []
        });
        await draft(fill(a.archivo, people), subject, to, eml);
    }

    const forum = config.correos.foro;
    const items: ForumItem[] = [];
    for (const m of matches) {
        const texts = await Promise.all(m.tasks.map((t) => doc.readTask(t.id)));
        items.push({
            requirement: m.requirement,
            functional: texts.map((t) => t?.functionalDesign ?? '').filter(Boolean).join(' '),
            technical: texts.map((t) => t?.technicalDesign ?? '').filter(Boolean).join(' '),
            type: m.type
        });
    }
    // Stories without requirement are described too, under their Azure DevOps ID.
    for (const t of unmatched) {
        const text = await doc.readTask(t.id);
        items.push({
            requirement: t.id,
            functional: text?.functionalDesign ?? '',
            technical: text?.technicalDesign ?? '',
            type: workTypeLabel(config, summaries.get(t.id)?.workItemType ?? ''),
            withoutRequirement: true
        });
    }
    const forumPeople = { ...people, especialista: forum.especialista };
    const forumSubject = fill(forum.asunto, forumPeople);
    const forumTo = recipients(forum.para);
    const forumEml = buildEml({
        to: forumTo,
        cc: recipients(forum.cc),
        subject: forumSubject,
        html: forumHtml(templates, {
            greeting: fill(forum.saludo, forumPeople),
            system: config.sistema,
            change: names.cambio,
            items,
            status: forum.estatusGeir,
            date: people.fechaLarga,
            time: forum.hora,
            engineers: fill(forum.ingenieroResponsable, forumPeople),
            specialist: forum.especialista,
            specialistLabel: forum.etiquetaEspecialista || 'Especialista',
            owner: forum.usuarioDueno,
            certifiers: forum.usuarioCertifica,
            backend: forum.recursosBackend ?? [],
            frontend: forum.recursosFrontend ?? [],
            signature
        }),
        inline
    });
    await draft(fill(forum.archivo, forumPeople), forumSubject, forumTo, forumEml);

    return {
        names,
        matches,
        warnings: [...new Set(warnings)],
        plan: {
            rows: plan.rows,
            totalDays: plan.totalDays,
            first: firstStory ? { id: firstStory.id, date: firstDate } : null,
            last: lastStory ? { id: lastStory.id, date: lastDate } : null,
            release
        },
        files,
        drafts,
        document: {
            location: describeLocation(location),
            name: stored.name,
            webUrl: stored.webUrl,
            created: false,
            tasks: await doc.listTasks()
        }
    };
}

/** A generated file, to download it from the page. */
export function adminFile(session: Session, name: string): Buffer | null {
    return session.adminFiles.get(name) ?? null;
}

/** Azure DevOps summaries (title, type, dates, parent) for every ID, once each. */
async function summariesFor(session: Session, ids: string[], warnings: string[]): Promise<Map<string, WorkItemSummary>> {
    const result = new Map<string, WorkItemSummary>();
    const unique = [...new Set(ids)];
    if (MOCK) {
        unique.forEach((id) => result.set(id, mockSummary(id)));
        return result;
    }
    const client = await azure(session);
    for (const id of unique) {
        try {
            const summary = await client.getSummary(id);
            if (summary) {
                result.set(id, summary);
            } else {
                warnings.push(`Azure DevOps no encontró el ID ${id}.`);
            }
        } catch (error: any) {
            warnings.push(`No se pudo leer ${id} en Azure DevOps: ${error.message}`);
        }
    }
    return result;
}

/** Writes a file in the same folder as the GEIR (local folder or OneDrive). */
async function writeSibling(session: Session, location: DocumentLocation, name: string, data: Buffer) {
    if (location.kind === 'local') {
        writeLocal({ filePath: path.join(path.dirname(location.filePath), name) }, data);
        return;
    }
    await withMicrosoft(session, async (client) => {
        const { siteUrl, serverRelativePath } = await client.normalize(location);
        await client.write(
            { kind: 'sharepoint', siteUrl, serverRelativePath: path.posix.join(path.posix.dirname(serverRelativePath), name) },
            data
        );
    });
}

/** Default values shown in the admin card. */
export function adminDefaults() {
    const config = loadAdminConfig();
    return {
        temporal: config.valorTemporal || 'TEMPORAL',
        files: [
            config.archivos.geir,
            ...config.archivos.derivados.map((d) => d.archivo),
            config.archivos.sprint,
            ...config.correos.aprobaciones.map((a) => a.archivo),
            config.correos.foro.archivo
        ]
    };
}
