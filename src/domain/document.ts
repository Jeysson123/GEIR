import { XlsxPackage } from './xlsx/package';
import { Worksheet } from './xlsx/worksheet';
import { RangeRef, cellName, formatRange } from './xlsx/refs';

export const SHEETS = {
    description: 'Descripción Requerimiento',
    design: 'Diseño Solución',
    log: 'Bitacora Pruebas Técnicas',
    uat: 'Plan de Pruebas UAT'
} as const;

export const DESIGN_LABELS = ['Diseño Funcional', 'Diseño Técnico', 'Diseño Seguridad'];
export const UAT_RESULTS = ['Satiscatorio', 'Fallido', 'Parcial'];

export interface EvidenceImage {
    base64: string;
    extension: string;
}

export interface GeirTask {
    id: string;
    title: string;
    number: number | null;
    // requirement this task fulfils (e.g. 2026-REQ-0101), '' when not related yet.
    requirement?: string;
    // Descripción Requerimiento
    benefit: string;
    // Diseño Solución
    functionalDesign: string;
    technicalDesign: string;
    securityDesign: string;
    // Bitácora Pruebas Técnicas
    workItemUrl: string;
    evidence: EvidenceImage | null;
    // Plan de Pruebas UAT
    testCase: string;
    given: string;
    when: string;
    then: string;
    result: string;
    requesterComments: string;
    evidenceLinkText: string;
    evidenceLinkUrl: string;
}

export interface TaskSummary {
    id: string;
    title: string;
    number: number | null;
    requirement: string;
}

/** One activity of the "Plan de Trabajo" table (Diseño Solución). */
export interface PlanRow {
    activity: string;
    days: number; // 0 leaves the row empty
    start: string; // YYYY-MM-DD
    end: string;
}

/** requirement code, e.g. 2026-REQ-0101 or 2026-REQ-0201-25. */
export const REQUIREMENT = /\b\d{4}-[A-Za-z]+-\d+(?:-\d+)?\b/;

/**
 * "7 - : 100007 : Título", ": 100007: Título", and with the requirement in front:
 * "1 - 2026-REQ-0103: 100004 - Título" or "2026-REQ-0101: 100002 - Título".
 */
const TASK_LINE = /^\s*(?:(\d+)\s*-\s*)?(?:(\d{4}-[A-Za-z]+-\d+(?:-\d+)?)\s*)?:?\s*(\d{3,})\s*[:\-–]\s*(.*)$/s;

export function parseTaskLine(
    text: string
): { number: number | null; requirement: string; id: string; title: string } | null {
    const m = TASK_LINE.exec(text.trim());
    if (!m) {
        return null;
    }
    return { number: m[1] ? Number(m[1]) : null, requirement: m[2] ?? '', id: m[3], title: m[4].trim() };
}

/** Line in "Yo Quiero": keeps the original format until a requirement is related. */
/**
 * "N - REQ: ID - title". `requirement: null` means "known to have none" (admin options):
 * "N - ID - title". Undefined keeps the form's original "N - : ID : title".
 */
export function descriptionLine(number: number, task: { id: string; title: string; requirement?: string | null }): string {
    if (task.requirement === null) {
        return `${number} - ${task.id} - ${task.title}`;
    }
    return task.requirement
        ? `${number} - ${task.requirement}: ${task.id} - ${task.title}`
        : `${number} - : ${task.id} : ${task.title}`;
}

/** Name of the task block in the other sheets. */
export function taskLabel(task: { id: string; title: string; requirement?: string | null }): string {
    if (task.requirement === null) {
        return `${task.id} - ${task.title}`;
    }
    return task.requirement ? `${task.requirement}: ${task.id} - ${task.title}` : `: ${task.id}: ${task.title}`;
}

function containsId(text: string, id: string): boolean {
    return new RegExp(`(^|\\D)${id}(\\D|$)`).test(text);
}

function startsWith(text: string, prefix: string): boolean {
    const norm = (s: string) =>
        s.normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase();
    return norm(text).startsWith(norm(prefix));
}

interface Block {
    top: number;
    bottom: number;
}

/**
 * The GEIR workbook: one row per task in each list of "Descripción Requerimiento" and
 * one block per task in the other three sheets. New tasks take the next free slot and,
 * when there is none, a new block copied from the last one so styles and merges match.
 */
export class GeirDocument {
    private constructor(readonly pkg: XlsxPackage) {}

    static async load(data: Buffer): Promise<GeirDocument> {
        const pkg = await XlsxPackage.load(data);
        const doc = new GeirDocument(pkg);
        // Fail early when the file is not a GEIR form.
        for (const name of Object.values(SHEETS)) {
            await pkg.sheet(name);
        }
        return doc;
    }

    save(): Promise<Buffer> {
        return this.pkg.save();
    }

    private sheet(name: string): Promise<Worksheet> {
        return this.pkg.sheet(name);
    }

    // ================================================================ lookup

    async listTasks(): Promise<TaskSummary[]> {
        const ws = await this.sheet(SHEETS.description);
        const { wantRows } = await this.descriptionLayout(ws);
        const tasks: TaskSummary[] = [];
        for (const row of wantRows) {
            const parsed = parseTaskLine(await ws.displayText(`B${row}`));
            if (parsed) {
                tasks.push({ id: parsed.id, title: parsed.title, number: parsed.number, requirement: parsed.requirement });
            }
        }
        // Tasks that only appear in the other sheets still count.
        for (const name of [SHEETS.design, SHEETS.log, SHEETS.uat]) {
            const sheet = await this.sheet(name);
            for (const block of await this.blocksFor(name, sheet)) {
                const parsed = parseTaskLine(await sheet.displayText(`A${block.top}`));
                if (parsed && !tasks.some((t) => t.id === parsed.id)) {
                    tasks.push({ id: parsed.id, title: parsed.title, number: parsed.number, requirement: parsed.requirement });
                }
            }
        }
        return tasks;
    }

    async hasTask(id: string): Promise<boolean> {
        return (await this.listTasks()).some((t) => t.id === id);
    }

    async readTask(id: string): Promise<GeirTask | null> {
        if (!(await this.hasTask(id))) {
            return null;
        }
        const task = emptyTask(id);

        const desc = await this.sheet(SHEETS.description);
        const layout = await this.descriptionLayout(desc);
        const wantRow = await this.findRow(desc, layout.wantRows, 'B', id);
        if (wantRow) {
            const parsed = parseTaskLine(await desc.displayText(`B${wantRow}`));
            if (parsed) {
                task.title = parsed.title;
                task.number = parsed.number;
                task.requirement = parsed.requirement;
            }
            const benefitRow = await this.findBenefitRow(desc, layout.benefitRows, layout.wantRows, wantRow, task.number);
            if (benefitRow) {
                task.benefit = stripNumber(await desc.displayText(`B${benefitRow}`));
            }
        }

        const design = await this.sheet(SHEETS.design);
        const designBlock = await this.findBlock(SHEETS.design, design, id);
        if (designBlock) {
            const line = parseTaskLine(await design.displayText(`A${designBlock.top}`));
            task.title = task.title || line?.title || '';
            task.requirement = task.requirement || line?.requirement || '';
            task.functionalDesign = await design.displayText(`C${designBlock.top}`);
            task.technicalDesign = await design.displayText(`C${designBlock.top + 1}`);
            task.securityDesign = await design.displayText(`C${designBlock.top + 2}`);
        }

        const log = await this.sheet(SHEETS.log);
        const logBlock = await this.findBlock(SHEETS.log, log, id);
        if (logBlock) {
            task.title = task.title || parseTaskLine(await log.displayText(`A${logBlock.top}`))?.title || '';
            task.workItemUrl =
                (await log.hyperlinkAt(`D${logBlock.bottom}`)) ?? (await log.displayText(`D${logBlock.bottom}`));
            const image =
                (await log.cellImage(`D${logBlock.top}`)) ??
                (await log.pictureIn({ top: logBlock.top, bottom: logBlock.bottom - 1, left: 1, right: 13 }));
            task.evidence = image ? { base64: image.data.toString('base64'), extension: image.extension } : null;
        }

        const uat = await this.sheet(SHEETS.uat);
        const uatBlock = await this.findBlock(SHEETS.uat, uat, id);
        if (uatBlock) {
            const r = uatBlock.top;
            task.title = task.title || parseTaskLine(await uat.displayText(`A${r}`))?.title || '';
            task.testCase = await uat.displayText(`B${r}`);
            task.given = await uat.displayText(`C${r}`);
            task.when = await uat.displayText(`D${r}`);
            task.then = await uat.displayText(`E${r}`);
            task.result = await uat.displayText(`F${r}`);
            task.requesterComments = await uat.displayText(`G${r}`);
            task.evidenceLinkText = await uat.displayText(`H${r}`);
            task.evidenceLinkUrl = (await uat.hyperlinkAt(`H${r}`)) ?? '';
        }
        return task;
    }

    // ================================================================ write

    /**
     * Adds the task in the next free space of every sheet, or updates it in place when
     * the ID is already there. `evidence: null` keeps whatever picture the file has.
     */
    async upsertTask(task: GeirTask): Promise<'added' | 'updated'> {
        const existed = await this.hasTask(task.id);
        const number = await this.writeDescription(task);
        await this.writeDesign(task);
        await this.writeLog(task);
        await this.writeUat(task);
        task.number = number;
        return existed ? 'updated' : 'added';
    }

    private async writeDescription(task: GeirTask): Promise<number> {
        const ws = await this.sheet(SHEETS.description);
        let layout = await this.descriptionLayout(ws);

        let wantRow = await this.findRow(ws, layout.wantRows, 'B', task.id);
        let number: number;
        let benefitRow: number | null = null;

        if (wantRow) {
            const parsed = parseTaskLine(await ws.displayText(`B${wantRow}`));
            const current = parsed?.number ?? task.number;
            benefitRow = await this.findBenefitRow(ws, layout.benefitRows, layout.wantRows, wantRow, current);
            number = current ?? (await this.nextNumber(ws, layout.wantRows));
        } else {
            number = await this.nextNumber(ws, layout.wantRows);
            wantRow = await this.nextListRow(ws, layout.wantHeader, layout.wantRows, layout.wantEnd);
            layout = await this.descriptionLayout(ws);
        }
        ws.setText(`B${wantRow}`, descriptionLine(number, task));

        if (!benefitRow) {
            benefitRow = await this.nextListRow(ws, layout.benefitHeader, layout.benefitRows, layout.benefitEnd);
        }
        ws.setText(`B${benefitRow}`, `${number} - ${task.benefit}`);
        ws.fitRowHeight(benefitRow, `${number} - ${task.benefit}`, 2, 10);
        await this.tidyDescription();
        return number;
    }

    private async writeDesign(task: GeirTask) {
        const ws = await this.sheet(SHEETS.design);
        const block = await this.blockForWrite(SHEETS.design, ws, task.id);
        const r = block.top;
        ws.setText(`A${r}`, taskLabel(task));
        DESIGN_LABELS.forEach((label, i) => ws.setText(`B${r + i}`, label));
        ws.setText(`C${r}`, task.functionalDesign);
        ws.setText(`C${r + 1}`, task.technicalDesign);
        ws.setText(`C${r + 2}`, task.securityDesign || 'N/A');
        // C:J is one merged cell per row.
        ws.fitRowHeight(r, task.functionalDesign, 3, 10);
        ws.fitRowHeight(r + 1, task.technicalDesign, 3, 10);
    }

    private async writeLog(task: GeirTask) {
        const ws = await this.sheet(SHEETS.log);
        const block = await this.blockForWrite(SHEETS.log, ws, task.id);
        ws.setText(`A${block.top}`, taskLabel(task));
        if (task.evidence) {
            // The picture area is the D:M merge at the top of the block (above the link row).
            const area = ws.mergeContaining(`D${block.top}`) ?? {
                top: block.top, bottom: block.bottom - 1, left: 4, right: 13
            };
            const pictureArea = { ...area, bottom: Math.min(area.bottom, block.bottom - 1) };
            await ws.removeAnchorsIn({ ...pictureArea, bottom: block.bottom });
            if (await ws.cellImage(`D${block.top}`)) {
                ws.setText(`D${block.top}`, '');
            }
            await ws.placePicture(pictureArea, Buffer.from(task.evidence.base64, 'base64'), task.evidence.extension);
        }
        if (task.workItemUrl) {
            const linkRange = ws.mergeContaining(`D${block.bottom}`) ?? {
                top: block.bottom, bottom: block.bottom, left: 4, right: 4
            };
            ws.setText(`D${block.bottom}`, task.workItemUrl);
            await ws.setHyperlink(formatRange(linkRange), task.workItemUrl);
        }
    }

    private async writeUat(task: GeirTask) {
        const ws = await this.sheet(SHEETS.uat);
        const block = await this.blockForWrite(SHEETS.uat, ws, task.id);
        const r = block.top;
        ws.setText(`A${r}`, taskLabel(task));
        ws.setText(`B${r}`, task.testCase);
        ws.setText(`C${r}`, task.given);
        ws.setText(`D${r}`, task.when);
        ws.setText(`E${r}`, task.then);
        ws.setText(`F${r}`, task.result);
        ws.setText(`G${r}`, task.requesterComments);
        const linkUrl = task.evidenceLinkUrl || task.workItemUrl;
        ws.setText(`H${r}`, task.evidenceLinkText || linkUrl);
        if (linkUrl) {
            await ws.setHyperlink(formatRange({ top: r, bottom: block.bottom, left: 8, right: 8 }), linkUrl, task.evidenceLinkText || linkUrl);
        }
    }

    // ================================================================ admin options

    /** "ID del Pedido:" in Descripción Requerimiento, e.g. "2026-REQ-0101, 2026-REQ-0102". */
    async setRequestIds(requirements: string[]) {
        const ws = await this.sheet(SHEETS.description);
        const row = await this.labelRow(ws, 'ID del Pedido');
        if (!row) {
            throw new Error('No se encontró "ID del Pedido" en Descripción Requerimiento');
        }
        const text = requirements.join(', ');
        ws.setText(`C${row}`, text);
        // C:J merged; at least two lines tall, like the original form.
        ws.fitRowHeight(row, text, 3, 10, 16, { shrink: true, min: 31.5 });
    }

    /**
     * Gives every line of the "Yo Quiero" / "Para Que" lists the format of the first line of
     * its list (style, height and the B:J merge), as in the original form. Lines added in
     * spare rows of a template could otherwise be left unmerged or with another height.
     */
    async tidyDescription() {
        const ws = await this.sheet(SHEETS.description);
        const layout = await this.descriptionLayout(ws);
        const lists: [number, number[], boolean][] = [
            [layout.wantHeader, layout.wantRows, false],
            [layout.benefitHeader, layout.benefitRows, true]
        ];
        for (const [header, rows, fit] of lists) {
            const template = header + 1;
            const merge = ws.mergeContaining(`B${template}`);
            for (const r of rows) {
                if (r <= template) {
                    continue;
                }
                const text = await ws.displayText(`B${r}`);
                if (!text) {
                    continue;
                }
                ws.copyRowFormat(template, r);
                if (merge && merge.top === merge.bottom && !ws.mergeContaining(`B${r}`)) {
                    let free = true;
                    for (let col = merge.left + 1; col <= merge.right; col++) {
                        if (ws.mergeContaining(cellName(col, r)) || (await ws.text(cellName(col, r)))) {
                            free = false;
                        }
                    }
                    if (free) {
                        ws.addMerge(formatRange({ ...merge, top: r, bottom: r }));
                    }
                }
                if (fit) {
                    ws.fitRowHeight(r, text, 2, 10);
                }
            }
        }
    }

    /** "Descripción:" mentions the sprint ("… solicitadas en Sprint 24 2025"). */
    async setSprint(sprint: string) {
        const ws = await this.sheet(SHEETS.description);
        const row = await this.labelRow(ws, 'Descripción');
        if (!row) {
            return;
        }
        const text = await ws.displayText(`C${row}`);
        if (/\bSprint\b/i.test(text)) {
            ws.setText(`C${row}`, text.replace(/\bSprint\b.*$/i, `Sprint ${sprint}`));
        }
    }

    /**
     * Writes the requirement in front of the task in the four sheets
     * ("2026-REQ-0101: 100002 - Título"). Returns false when the task is not in the file.
     */
    async relateRequirement(id: string, requirement: string | null): Promise<boolean> {
        if (!(await this.hasTask(id))) {
            return false;
        }
        const desc = await this.sheet(SHEETS.description);
        const layout = await this.descriptionLayout(desc);
        const wantRow = await this.findRow(desc, layout.wantRows, 'B', id);
        if (wantRow) {
            const line = parseTaskLine(await desc.displayText(`B${wantRow}`));
            if (line?.number !== null && line?.number !== undefined) {
                desc.setText(`B${wantRow}`, descriptionLine(line.number, { id, title: line.title, requirement }));
            }
        }
        for (const name of [SHEETS.design, SHEETS.log, SHEETS.uat]) {
            const ws = await this.sheet(name);
            const block = await this.findBlock(name, ws, id);
            const line = block ? parseTaskLine(await ws.displayText(`A${block.top}`)) : null;
            if (block && line) {
                ws.setText(`A${block.top}`, taskLabel({ id, title: line.title, requirement }));
            }
        }
        return true;
    }

    /** Fills "Plan de Trabajo" (Diseño Solución): days and dates per activity. */
    async setWorkPlan(rows: PlanRow[]): Promise<string[]> {
        const ws = await this.sheet(SHEETS.design);
        const header = await this.labelRow(ws, 'Plan de Trabajo');
        if (!header) {
            throw new Error('No se encontró "Plan de Trabajo" en Diseño Solución');
        }
        const missing: string[] = [];
        for (const row of rows) {
            let found = 0;
            for (let r = header + 1; r <= Math.min(ws.maxRow, header + 30); r++) {
                const label = (await ws.displayText(`A${r}`)).replace(/\*/g, '');
                if (label && startsWith(label, row.activity.replace(/\*/g, '')) && startsWith(row.activity.replace(/\*/g, ''), label)) {
                    found = r;
                    break;
                }
                if (startsWith(label, 'Total')) {
                    break;
                }
            }
            if (!found) {
                missing.push(row.activity);
                continue;
            }
            if (row.days > 0) {
                ws.setNumber(`C${found}`, row.days);
                ws.setNumber(`D${found}`, excelDate(row.start));
                ws.setNumber(`E${found}`, excelDate(row.end));
            } else {
                ['C', 'D', 'E'].forEach((col) => ws.setText(`${col}${found}`, ''));
            }
        }
        // "Total días estimados" is a SUM of the days: let Excel recompute it.
        await this.pkg.recalculateOnOpen();
        return missing;
    }

    private async labelRow(ws: Worksheet, label: string): Promise<number> {
        for (let r = 1; r <= ws.maxRow; r++) {
            if (startsWith(await ws.displayText(`A${r}`), label)) {
                return r;
            }
        }
        return 0;
    }

    // ================================================================ delete

    /**
     * Removes the task from the four sheets. Its rows/blocks are deleted (the last one of a
     * list or sheet is emptied instead, so the form keeps a slot to copy from) and the tasks
     * numbered after it in "Descripción Requerimiento" move up one number.
     */
    async deleteTask(id: string): Promise<boolean> {
        if (!(await this.hasTask(id))) {
            return false;
        }
        await this.deleteFromDescription(id);
        for (const name of [SHEETS.design, SHEETS.log, SHEETS.uat]) {
            const ws = await this.sheet(name);
            const block = await this.findBlock(name, ws, id);
            if (!block) {
                continue;
            }
            if ((await this.blocksFor(name, ws)).length > 1) {
                await ws.deleteRows(block.top, block.bottom - block.top + 1);
            } else {
                await this.clearBlock(name, ws, block);
            }
        }
        return true;
    }

    private async deleteFromDescription(id: string) {
        const ws = await this.sheet(SHEETS.description);
        const layout = await this.descriptionLayout(ws);
        const wantRow = await this.findRow(ws, layout.wantRows, 'B', id);
        if (!wantRow) {
            return;
        }
        const number = parseTaskLine(await ws.displayText(`B${wantRow}`))?.number ?? null;
        const benefitRow = await this.findBenefitRow(ws, layout.benefitRows, layout.wantRows, wantRow, number);

        // Benefit first: it is below, so deleting it does not move the "Yo Quiero" row.
        if (benefitRow) {
            await this.removeListRow(ws, benefitRow, layout.benefitHeader, layout.benefitEnd);
        }
        await this.removeListRow(ws, wantRow, layout.wantHeader, layout.wantEnd);

        if (number !== null) {
            await this.renumberAfter(ws, number);
        }
    }

    /** Deletes a list row, or empties it when it is the list's only item row (header + item + spacer). */
    private async removeListRow(ws: Worksheet, row: number, header: number, end: number) {
        const itemRows = end - header - 2; // rows between the header and the spacer row
        if (itemRows > 1) {
            await ws.deleteRows(row, 1);
        } else {
            ws.setText(`B${row}`, '');
        }
    }

    /** "8 - …" becomes "7 - …" in both lists once task number 7 is gone. */
    private async renumberAfter(ws: Worksheet, removed: number) {
        const layout = await this.descriptionLayout(ws);
        for (const r of [...layout.wantRows.slice(1), ...layout.benefitRows.slice(1)]) {
            const text = await ws.displayText(`B${r}`);
            const m = /^(\s*)(\d+)(\s*-)/.exec(text);
            if (m && Number(m[2]) > removed) {
                ws.setText(`B${r}`, `${m[1]}${Number(m[2]) - 1}${m[3]}${text.slice(m[0].length)}`);
            }
        }
    }

    // ================================================================ Descripción Requerimiento

    private async descriptionLayout(ws: Worksheet) {
        let wantHeader = 0;
        let benefitHeader = 0;
        let afterBenefit = 0;
        for (let r = 1; r <= ws.maxRow; r++) {
            const a = await ws.displayText(`A${r}`);
            if (!a) {
                continue;
            }
            if (!wantHeader && startsWith(a, 'Yo Quiero')) {
                wantHeader = r;
            } else if (wantHeader && !benefitHeader && startsWith(a, 'Para Que')) {
                benefitHeader = r;
            } else if (benefitHeader && !afterBenefit && r > benefitHeader) {
                afterBenefit = r;
            }
        }
        if (!wantHeader || !benefitHeader) {
            throw new Error('No se encontraron las secciones "Yo Quiero" / "Para Que" en Descripción Requerimiento');
        }
        const wantEnd = benefitHeader; // exclusive
        const benefitEnd = afterBenefit || ws.maxRow + 1;
        const range = (from: number, to: number) => Array.from({ length: Math.max(0, to - from) }, (_, i) => from + i);
        return {
            wantHeader,
            benefitHeader,
            wantEnd,
            benefitEnd,
            wantRows: range(wantHeader, wantEnd),
            benefitRows: range(benefitHeader, benefitEnd)
        };
    }

    private async findRow(ws: Worksheet, rows: number[], col: string, id: string): Promise<number | null> {
        for (const r of rows) {
            const text = await ws.displayText(`${col}${r}`);
            const parsed = parseTaskLine(text);
            if (parsed?.id === id || (parsed === null && text && containsId(text, id) && /:\s*\d/.test(text))) {
                return r;
            }
        }
        return null;
    }

    private async findBenefitRow(
        ws: Worksheet,
        benefitRows: number[],
        wantRows: number[],
        wantRow: number,
        number: number | null
    ): Promise<number | null> {
        if (number !== null) {
            for (const r of benefitRows.slice(1)) {
                const m = /^\s*(\d+)\s*-/.exec(await ws.displayText(`B${r}`));
                if (m && Number(m[1]) === number) {
                    return r;
                }
            }
        }
        // Fall back to the same position in the list.
        const filledWants: number[] = [];
        for (const r of wantRows.slice(1)) {
            if (await ws.displayText(`B${r}`)) {
                filledWants.push(r);
            }
        }
        const position = filledWants.indexOf(wantRow);
        const filledBenefits: number[] = [];
        for (const r of benefitRows.slice(1)) {
            if (await ws.displayText(`B${r}`)) {
                filledBenefits.push(r);
            }
        }
        return position >= 0 ? filledBenefits[position] ?? null : null;
    }

    private async nextNumber(ws: Worksheet, rows: number[]): Promise<number> {
        let max = 0;
        for (const r of rows) {
            const m = /^\s*(\d+)\s*-/.exec(await ws.displayText(`B${r}`));
            if (m) {
                max = Math.max(max, Number(m[1]));
            }
        }
        return max + 1;
    }

    /**
     * Next empty row under a list header. Keeps one blank spacer row before the next
     * section; when the list is full a row is inserted with the previous row's format.
     */
    private async nextListRow(ws: Worksheet, header: number, rows: number[], end: number): Promise<number> {
        let lastFilled = header;
        for (const r of rows) {
            if (r > header && (await ws.displayText(`B${r}`))) {
                lastFilled = r;
            }
        }
        const target = lastFilled + 1;
        if (target < end - 1) {
            return target;
        }
        const template = lastFilled > header ? lastFilled : Math.min(header + 1, end - 1);
        await ws.insertRows(target, 1);
        ws.copyRowFormat(template, target);
        const merge = ws.mergeContaining(`B${template}`);
        if (merge && merge.top === merge.bottom) {
            ws.addMerge(formatRange({ ...merge, top: target, bottom: target }));
        }
        return target;
    }

    // ================================================================ block sheets

    private async headerRow(name: string, ws: Worksheet): Promise<number> {
        const marker = name === SHEETS.design ? 'Entregables' : name === SHEETS.log ? 'Bitácora' : 'Funcionalidad';
        for (let r = 1; r <= ws.maxRow; r++) {
            if (startsWith(await ws.displayText(`A${r}`), marker)) {
                return r;
            }
        }
        throw new Error(`No se encontró el encabezado "${marker}" en ${name}`);
    }

    /** Task blocks in sheet order. Each block is identified by the merge of its A cell. */
    async blocksFor(name: string, ws: Worksheet): Promise<Block[]> {
        const header = await this.headerRow(name, ws);
        const merges = ws.merges();

        if (name === SHEETS.log) {
            const starts = merges
                .filter((m) => m.left === 1 && m.right === 3 && m.top > header && m.bottom - m.top >= 3)
                .sort((a, b) => a.top - b.top);
            return starts.map((m) => {
                // The link row is the last single-row D:M merge at or right after the name merge.
                const links = merges.filter(
                    (x) => x.left === 4 && x.top === x.bottom && x.top >= m.top && x.top <= m.bottom + 1
                );
                const bottom = links.length ? Math.max(...links.map((x) => x.top)) : m.bottom + 1;
                return { top: m.top, bottom };
            });
        }

        const footer = name === SHEETS.design ? 'El componente modificado' : 'Dado';
        let footerRow = ws.maxRow + 1;
        for (let r = header + 1; r <= ws.maxRow; r++) {
            const a = await ws.displayText(`A${r}`);
            if (a && startsWith(a, footer) && !ws.mergeContaining(`A${r}`)) {
                footerRow = r;
                break;
            }
        }
        return merges
            .filter((m) => m.left === 1 && m.right === 1 && m.top > header && m.bottom < footerRow && m.bottom - m.top === 2)
            .sort((a, b) => a.top - b.top)
            .map((m) => ({ top: m.top, bottom: m.bottom }));
    }

    private async findBlock(name: string, ws: Worksheet, id: string): Promise<Block | null> {
        for (const block of await this.blocksFor(name, ws)) {
            const text = await ws.displayText(`A${block.top}`);
            if (parseTaskLine(text)?.id === id || containsId(text, id)) {
                return block;
            }
        }
        return null;
    }

    private async blockForWrite(name: string, ws: Worksheet, id: string): Promise<Block> {
        const existing = await this.findBlock(name, ws, id);
        if (existing) {
            return existing;
        }
        const blocks = await this.blocksFor(name, ws);
        for (const block of blocks) {
            if (!(await ws.displayText(`A${block.top}`))) {
                return block; // empty slot
            }
        }
        if (!blocks.length) {
            throw new Error(`La hoja ${name} no tiene bloques de ejemplo para copiar el formato`);
        }
        // Prefer a standard-sized block as the format template (Bitácora has some double ones).
        const standardSize = name === SHEETS.log ? 5 : 3;
        const template =
            [...blocks].reverse().find((b) => b.bottom - b.top + 1 === standardSize) ?? blocks[blocks.length - 1];
        const last = blocks[blocks.length - 1];
        return this.cloneBlock(ws, template, last.bottom + 1);
    }

    private async cloneBlock(ws: Worksheet, template: Block, at: number): Promise<Block> {
        const size = template.bottom - template.top + 1;
        const templateMerges = ws
            .merges()
            .filter((m) => m.top >= template.top && m.bottom <= template.bottom);
        await ws.insertRows(at, size);
        const offset = at - template.top;
        for (let i = 0; i < size; i++) {
            ws.copyRowFormat(template.top + i, at + i);
        }
        for (const m of templateMerges) {
            ws.addMerge(formatRange({ ...m, top: m.top + offset, bottom: m.bottom + offset }));
        }
        return { top: at, bottom: at + size - 1 };
    }

    // ================================================================ blank template

    /** Turns a filled GEIR into an empty form with one free slot per section. */
    async clearAllTasks() {
        const desc = await this.sheet(SHEETS.description);
        for (const which of ['want', 'benefit'] as const) {
            const layout = await this.descriptionLayout(desc);
            const header = which === 'want' ? layout.wantHeader : layout.benefitHeader;
            const end = which === 'want' ? layout.wantEnd : layout.benefitEnd;
            // Keep the first item row and one spacer row.
            const extra = end - header - 3;
            if (extra > 0) {
                await desc.deleteRows(header + 2, extra);
            }
            desc.setText(`B${header + 1}`, '');
        }

        for (const name of [SHEETS.design, SHEETS.log, SHEETS.uat]) {
            const ws = await this.sheet(name);
            const blocks = await this.blocksFor(name, ws);
            const standardSize = name === SHEETS.log ? 5 : 3;
            const keep =
                [...blocks].reverse().find((b) => b.bottom - b.top + 1 === standardSize) ?? blocks[blocks.length - 1];
            for (const block of [...blocks].reverse()) {
                if (block !== keep) {
                    await ws.deleteRows(block.top, block.bottom - block.top + 1);
                }
            }
            const kept = (await this.blocksFor(name, ws))[0];
            await this.clearBlock(name, ws, kept);
        }
    }

    private async clearBlock(name: string, ws: Worksheet, block: Block) {
        const area: RangeRef = { top: block.top, bottom: block.bottom, left: 1, right: name === SHEETS.log ? 13 : 10 };
        await ws.removeHyperlinksIn(area);
        await ws.removeAnchorsIn(area);
        for (let r = block.top; r <= block.bottom; r++) {
            for (let c = 1; c <= area.right; c++) {
                const ref = cellName(c, r);
                const isLabel = name === SHEETS.design && c === 2;
                if (!isLabel && ((await ws.displayText(ref)) || (await ws.cellImage(ref)))) {
                    ws.setText(ref, '');
                }
            }
        }
        if (name === SHEETS.design) {
            ws.setText(`C${block.top + 2}`, 'N/A');
        }
    }
}

export function emptyTask(id: string): GeirTask {
    return {
        id,
        title: '',
        number: null,
        requirement: '',
        benefit: '',
        functionalDesign: '',
        technicalDesign: '',
        securityDesign: 'N/A',
        workItemUrl: '',
        evidence: null,
        testCase: '',
        given: '',
        when: '',
        then: '',
        result: UAT_RESULTS[0],
        requesterComments: '',
        evidenceLinkText: '',
        evidenceLinkUrl: ''
    };
}

/** Excel serial date (days since 1899-12-30) for YYYY-MM-DD. */
function excelDate(iso: string): number {
    const [y, m, d] = iso.split('-').map(Number);
    return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
}

function stripNumber(text: string): string {
    return text.replace(/^\s*\d+\s*-\s*/, '').trim();
}
