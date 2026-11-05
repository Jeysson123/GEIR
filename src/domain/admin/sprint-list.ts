import ExcelJS from 'exceljs';

export interface SprintStory {
    requirement: string; // empty when the story has none
    id: string;
    title: string;
    description: string; // "ID: título", the same text as the e-mail table
    type: string; // "Funcionalidad" / "Corrección Defecto"
    assignedTo: string;
}

// Layout of the sprint list template (config/plantillas/sprint.xlsx).
const GENERAL = { sheet: 'Sheet1', header: 3, first: 4, linesGap: 11 }; // B:G table, lines in E
const MAIL = { sheet: 'Sheet2', header: 4, first: 5 }; // B:D table

/**
 * The sprint list, written into the original workbook so fonts, colors, borders and column
 * widths stay exactly as in the template. Every story of the GEIR is listed, with or
 * without a requirement. Sheet1: "Información general" and the "REQ: ID - Descripción"
 * lines; Sheet2: the table used in the notification e-mails.
 */
export async function buildSprintWorkbook(template: Buffer, stories: SprintStory[]): Promise<Buffer> {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(template as any);

    const general = sheet(wb, GENERAL.sheet);
    const generalStyle = rowStyles(general, GENERAL.first, 1, 7);
    clearFrom(general, GENERAL.first);
    stories.forEach((s, i) => {
        const r = GENERAL.first + i;
        write(general, r, generalStyle, [i + 1, s.requirement, Number(s.id) || s.id, s.title, `${s.id} - ${s.title}`, s.type, s.assignedTo]);
    });
    const linesRow = GENERAL.first + stories.length - 1 + GENERAL.linesGap;
    stories.forEach((s, i) => {
        general.getCell(`E${linesRow + i}`).value = s.requirement ? `${s.requirement}: ${s.id} - ${s.title}` : `${s.id} - ${s.title}`;
    });

    const mail = sheet(wb, MAIL.sheet);
    const mailStyle = rowStyles(mail, MAIL.first, 1, 4);
    clearFrom(mail, MAIL.first);
    const descriptionWidth = mail.getColumn(3).width || 75;
    stories.forEach((s, i) => {
        const r = MAIL.first + i;
        write(mail, r, mailStyle, [i + 1, s.requirement, s.description, s.type]);
        // Column C wraps: give long descriptions the lines they need, as in the template.
        const lines = Math.ceil((s.description.length * 1.1) / descriptionWidth);
        if (lines > 1) {
            mail.getRow(r).height = lines * 15;
        }
    });

    return Buffer.from(await wb.xlsx.writeBuffer());
}

function sheet(wb: ExcelJS.Workbook, name: string): ExcelJS.Worksheet {
    const ws = wb.getWorksheet(name) ?? wb.worksheets[name === GENERAL.sheet ? 0 : 1];
    if (!ws) {
        throw new Error(`La plantilla del sprint no tiene la hoja ${name}.`);
    }
    return ws;
}

/** Styles of the first data row (columns from..to), copied to every row written. */
function rowStyles(ws: ExcelJS.Worksheet, row: number, from: number, to: number): Partial<ExcelJS.Style>[] {
    const styles: Partial<ExcelJS.Style>[] = [];
    for (let c = from; c <= to; c++) {
        styles.push(JSON.parse(JSON.stringify(ws.getRow(row).getCell(c).style ?? {})));
    }
    return styles;
}

/** Empties the example rows of the template (values, styles and heights). */
function clearFrom(ws: ExcelJS.Worksheet, row: number) {
    for (let r = row; r <= ws.rowCount; r++) {
        const current = ws.getRow(r);
        current.eachCell({ includeEmpty: true }, (cell) => {
            cell.value = null;
            cell.style = {};
        });
        current.height = undefined as unknown as number;
    }
}

function write(ws: ExcelJS.Worksheet, r: number, styles: Partial<ExcelJS.Style>[], values: (string | number)[]) {
    const row = ws.getRow(r);
    values.forEach((value, i) => {
        const cell = row.getCell(i + 1);
        cell.style = JSON.parse(JSON.stringify(styles[i] ?? {}));
        cell.value = value === '' ? null : value;
    });
}
