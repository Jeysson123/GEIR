import * as path from 'path';
import { DOMParser } from '@xmldom/xmldom';
import { NS, REL_TYPE, XlsxPackage, resolvePartPath } from './package';
import {
    RangeRef,
    cellName,
    formatRange,
    parseCell,
    parseRange,
    shiftFormula,
    shiftRange,
    shiftSqref
} from './refs';

// Child order of <worksheet> required by the SpreadsheetML schema.
const WORKSHEET_ORDER = [
    'sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr',
    'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate',
    'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations',
    'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks',
    'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing',
    'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems',
    'tableParts', 'extLst'
];

const MAX_CELL_TEXT = 32767;

export interface CellImage {
    data: Buffer;
    extension: string;
}

function localName(el: Node): string {
    const name = el.nodeName;
    const i = name.indexOf(':');
    return i === -1 ? name : name.slice(i + 1);
}

function childElements(el: Element): Element[] {
    return Array.from(el.childNodes).filter((n): n is Element => n.nodeType === 1);
}

function firstChild(el: Element, name: string): Element | null {
    return childElements(el).find((c) => localName(c) === name) ?? null;
}

export class Worksheet {
    doc!: Document;
    private sheetData!: Element;
    private rows = new Map<number, Element>();

    constructor(
        private pkg: XlsxPackage,
        readonly part: string,
        readonly name: string,
        readonly index: number
    ) {}

    async init() {
        this.doc = await this.pkg.xml(this.part);
        const sheetData = this.doc.getElementsByTagName('sheetData')[0];
        if (!sheetData) {
            throw new Error(`Worksheet ${this.name} has no sheetData`);
        }
        this.sheetData = sheetData;
        this.indexRows();
    }

    private indexRows() {
        this.rows.clear();
        for (const row of childElements(this.sheetData)) {
            if (localName(row) === 'row') {
                this.rows.set(Number(row.getAttribute('r')), row);
            }
        }
    }

    private touch() {
        this.pkg.markDirty(this.part);
    }

    get maxRow(): number {
        return Math.max(0, ...this.rows.keys());
    }

    // ---------------------------------------------------------------- rows & cells

    private rowEl(r: number, create: boolean): Element | null {
        const existing = this.rows.get(r);
        if (existing || !create) {
            return existing ?? null;
        }
        const row = this.doc.createElementNS(NS.main, 'row');
        row.setAttribute('r', String(r));
        const next = [...this.rows.keys()].filter((k) => k > r).sort((a, b) => a - b)[0];
        this.sheetData.insertBefore(row, next !== undefined ? this.rows.get(next)! : null);
        this.rows.set(r, row);
        this.touch();
        return row;
    }

    private cellEl(ref: string, create: boolean): Element | null {
        const { col, row } = parseCell(ref);
        const rowEl = this.rowEl(row, create);
        if (!rowEl) {
            return null;
        }
        const name = cellName(col, row);
        let before: Element | null = null;
        for (const c of childElements(rowEl)) {
            if (localName(c) !== 'c') {
                continue;
            }
            const cref = c.getAttribute('r');
            if (cref === name) {
                return c;
            }
            if (cref && parseCell(cref).col > col && !before) {
                before = c;
            }
        }
        if (!create) {
            return null;
        }
        const c = this.doc.createElementNS(NS.main, 'c');
        c.setAttribute('r', name);
        // A new cell takes the row's default style, like Excel does.
        const rowStyle = rowEl.getAttribute('customFormat') === '1' ? rowEl.getAttribute('s') : null;
        if (rowStyle) {
            c.setAttribute('s', rowStyle);
        }
        rowEl.insertBefore(c, before);
        this.touch();
        return c;
    }

    async text(ref: string): Promise<string> {
        const c = this.cellEl(ref, false);
        if (!c) {
            return '';
        }
        const type = c.getAttribute('t');
        if (type === 'inlineStr') {
            const is = firstChild(c, 'is');
            return is
                ? Array.from(is.getElementsByTagName('t'))
                      .filter((t) => localName(t.parentNode as Node) !== 'rPh')
                      .map((t) => t.textContent ?? '')
                      .join('')
                : '';
        }
        const v = firstChild(c, 'v');
        const raw = v?.textContent ?? '';
        if (type === 's') {
            const strings = await this.pkg.sharedStrings();
            return strings[Number(raw)] ?? '';
        }
        if (type === 'e' && c.getAttribute('vm')) {
            return ''; // in-cell picture
        }
        return raw;
    }

    /** Text of a cell, or of the top-left cell of the merge that contains it. */
    async displayText(ref: string): Promise<string> {
        const merge = this.mergeContaining(ref);
        return this.text(merge ? cellName(merge.left, merge.top) : ref);
    }

    setText(ref: string, value: string | null | undefined) {
        const c = this.cellEl(ref, true)!;
        for (const child of childElements(c)) {
            c.removeChild(child);
        }
        c.removeAttribute('t');
        c.removeAttribute('vm');
        c.removeAttribute('cm');
        const text = (value ?? '').replace(/\r\n/g, '\n').slice(0, MAX_CELL_TEXT);
        if (text !== '') {
            c.setAttribute('t', 'inlineStr');
            const is = this.doc.createElementNS(NS.main, 'is');
            const t = this.doc.createElementNS(NS.main, 't');
            t.setAttribute('xml:space', 'preserve');
            t.appendChild(this.doc.createTextNode(text));
            is.appendChild(t);
            c.appendChild(is);
        }
        this.touch();
    }

    setNumber(ref: string, value: number) {
        const c = this.cellEl(ref, true)!;
        for (const child of childElements(c)) {
            c.removeChild(child);
        }
        c.removeAttribute('t');
        c.removeAttribute('vm');
        const v = this.doc.createElementNS(NS.main, 'v');
        v.appendChild(this.doc.createTextNode(String(value)));
        c.appendChild(v);
        this.touch();
    }

    style(ref: string): string | null {
        return this.cellEl(ref, false)?.getAttribute('s') ?? null;
    }

    /** Copies row height/style and every cell's style (not values) from one row to another. */
    copyRowFormat(fromRow: number, toRow: number) {
        const src = this.rowEl(fromRow, false);
        const dst = this.rowEl(toRow, true)!;
        if (!src) {
            return;
        }
        for (const attr of ['ht', 'customHeight', 's', 'customFormat', 'spans', 'thickBot', 'thickTop']) {
            const value = src.getAttribute(attr);
            if (value !== null) {
                dst.setAttribute(attr, value);
            } else {
                dst.removeAttribute(attr);
            }
        }
        for (const c of childElements(src)) {
            if (localName(c) !== 'c') {
                continue;
            }
            const { col } = parseCell(c.getAttribute('r') ?? '');
            const target = this.cellEl(cellName(col, toRow), true)!;
            const s = c.getAttribute('s');
            if (s !== null) {
                target.setAttribute('s', s);
            }
        }
        this.touch();
    }

    // ---------------------------------------------------------------- merges

    private mergeCellsEl(create: boolean): Element | null {
        let el = this.doc.getElementsByTagName('mergeCells')[0] ?? null;
        if (!el && create) {
            el = this.insertWorksheetChild('mergeCells');
        }
        return el;
    }

    merges(): RangeRef[] {
        const el = this.mergeCellsEl(false);
        if (!el) {
            return [];
        }
        return Array.from(el.getElementsByTagName('mergeCell')).map((m) => parseRange(m.getAttribute('ref') ?? ''));
    }

    mergeContaining(ref: string): RangeRef | null {
        const { col, row } = parseCell(ref);
        return (
            this.merges().find((m) => row >= m.top && row <= m.bottom && col >= m.left && col <= m.right) ?? null
        );
    }

    addMerge(ref: string) {
        const range = parseRange(ref);
        const overlaps = this.merges().some(
            (m) => !(m.bottom < range.top || m.top > range.bottom || m.right < range.left || m.left > range.right)
        );
        if (overlaps) {
            return;
        }
        const el = this.mergeCellsEl(true)!;
        const m = this.doc.createElementNS(NS.main, 'mergeCell');
        m.setAttribute('ref', formatRange(range));
        el.appendChild(m);
        el.setAttribute('count', String(el.getElementsByTagName('mergeCell').length));
        this.touch();
    }

    // ---------------------------------------------------------------- hyperlinks

    hyperlinkRefs(): { ref: string; rid: string | null; el: Element }[] {
        const container = this.doc.getElementsByTagName('hyperlinks')[0];
        if (!container) {
            return [];
        }
        return Array.from(container.getElementsByTagName('hyperlink')).map((el) => ({
            ref: el.getAttribute('ref') ?? '',
            rid: el.getAttributeNS(NS.rel, 'id') || el.getAttribute('r:id'),
            el
        }));
    }

    async hyperlinkAt(ref: string): Promise<string | null> {
        const { col, row } = parseCell(ref);
        const rels = await this.pkg.relationships(this.part);
        for (const link of this.hyperlinkRefs()) {
            const r = parseRange(link.ref);
            if (row >= r.top && row <= r.bottom && col >= r.left && col <= r.right && link.rid) {
                return rels.find((x) => x.id === link.rid)?.target ?? null;
            }
        }
        return null;
    }

    async removeHyperlinksIn(range: RangeRef) {
        for (const link of this.hyperlinkRefs()) {
            const r = parseRange(link.ref);
            const overlaps = !(r.bottom < range.top || r.top > range.bottom || r.right < range.left || r.left > range.right);
            if (overlaps) {
                link.el.parentNode?.removeChild(link.el);
                if (link.rid) {
                    await this.pkg.removeRelationship(this.part, link.rid);
                }
                this.touch();
            }
        }
        const container = this.doc.getElementsByTagName('hyperlinks')[0];
        if (container && container.getElementsByTagName('hyperlink').length === 0) {
            container.parentNode?.removeChild(container);
        }
    }

    async setHyperlink(ref: string, url: string, display?: string) {
        const range = parseRange(ref);
        await this.removeHyperlinksIn(range);
        const rid = await this.pkg.addRelationship(this.part, REL_TYPE.hyperlink, url, true);
        let container = this.doc.getElementsByTagName('hyperlinks')[0];
        if (!container) {
            container = this.insertWorksheetChild('hyperlinks');
        }
        const el = this.doc.createElementNS(NS.main, 'hyperlink');
        el.setAttribute('ref', formatRange(range));
        el.setAttributeNS(NS.rel, 'r:id', rid);
        el.setAttribute('display', display ?? url);
        container.appendChild(el);
        this.touch();
    }

    private insertWorksheetChild(name: string): Element {
        const root = this.doc.documentElement;
        const order = WORKSHEET_ORDER.indexOf(name);
        const before = childElements(root).find((c) => WORKSHEET_ORDER.indexOf(localName(c)) > order) ?? null;
        const el = this.doc.createElementNS(NS.main, name);
        root.insertBefore(el, before);
        this.touch();
        return el;
    }

    // ---------------------------------------------------------------- row insertion / deletion

    async insertRows(at: number, count: number) {
        await this.shiftRows(at, count);
    }

    async deleteRows(at: number, count: number) {
        const range: RangeRef = { top: at, bottom: at + count - 1, left: 1, right: 16384 };
        await this.removeHyperlinksIn(range);
        await this.removeAnchorsIn(range);
        for (let r = at; r < at + count; r++) {
            const row = this.rows.get(r);
            if (row) {
                this.sheetData.removeChild(row);
                this.rows.delete(r);
            }
        }
        await this.shiftRows(at, -count);
    }

    private async shiftRows(at: number, delta: number) {
        const moveRow = (r: number) => (delta > 0 ? (r >= at ? r + delta : r) : r >= at - delta ? r + delta : r);

        // Cells and rows.
        for (const row of childElements(this.sheetData)) {
            const r = Number(row.getAttribute('r'));
            const nr = moveRow(r);
            if (nr === r) {
                continue;
            }
            row.setAttribute('r', String(nr));
            for (const c of childElements(row)) {
                const ref = c.getAttribute('r');
                if (ref) {
                    c.setAttribute('r', cellName(parseCell(ref).col, nr));
                }
            }
        }

        // Formulas anywhere in the sheet, including shared-formula ranges.
        for (const f of Array.from(this.doc.getElementsByTagName('f'))) {
            if (f.textContent) {
                f.textContent = shiftFormula(f.textContent, at, delta);
            }
            const ref = f.getAttribute('ref');
            if (ref) {
                f.setAttribute('ref', shiftSqref(ref, at, delta));
            }
        }

        // Merges.
        const mergeEl = this.mergeCellsEl(false);
        if (mergeEl) {
            for (const m of Array.from(mergeEl.getElementsByTagName('mergeCell'))) {
                const shifted = shiftRange(parseRange(m.getAttribute('ref') ?? ''), at, delta);
                if (shifted && (shifted.top !== shifted.bottom || shifted.left !== shifted.right)) {
                    m.setAttribute('ref', formatRange(shifted));
                } else {
                    mergeEl.removeChild(m);
                }
            }
            mergeEl.setAttribute('count', String(mergeEl.getElementsByTagName('mergeCell').length));
        }

        // Hyperlinks.
        for (const link of this.hyperlinkRefs()) {
            const shifted = shiftRange(parseRange(link.ref), at, delta);
            if (shifted) {
                link.el.setAttribute('ref', formatRange(shifted));
            }
        }

        // Every sqref attribute (validations, conditional formats, selections, protected
        // ranges) and the x14 extension equivalents stored as <xm:sqref> text.
        const all = Array.from(this.doc.getElementsByTagName('*'));
        for (const el of all) {
            const sqref = el.getAttribute('sqref');
            if (sqref) {
                const next = shiftSqref(sqref, at, delta);
                if (next) {
                    el.setAttribute('sqref', next);
                }
            }
            if (localName(el) === 'sqref' && el.textContent) {
                const next = shiftSqref(el.textContent, at, delta);
                if (next) {
                    el.textContent = next;
                }
            }
            if (localName(el) === 'selection' && el.getAttribute('activeCell')) {
                const shifted = shiftRange(parseRange(el.getAttribute('activeCell')!), at, delta);
                if (shifted) {
                    el.setAttribute('activeCell', formatRange(shifted));
                }
            }
            if (localName(el) === 'brk' && localName(el.parentNode as Node) === 'rowBreaks') {
                const id = Number(el.getAttribute('id'));
                el.setAttribute('id', String(moveRow(id)));
            }
        }

        // Dimension.
        const dim = this.doc.getElementsByTagName('dimension')[0];
        if (dim?.getAttribute('ref')) {
            const r = parseRange(dim.getAttribute('ref')!);
            r.bottom = Math.max(r.top, r.bottom + delta);
            dim.setAttribute('ref', formatRange(r));
        }

        await this.shiftAnchors(at, delta);
        await this.shiftDefinedNames(at, delta);
        await this.pkg.dropCalcChain();

        this.indexRows();
        this.touch();
    }

    private async drawingPart(): Promise<string | null> {
        const rels = await this.pkg.relationships(this.part);
        const rel = rels.find((r) => r.type === REL_TYPE.drawing);
        return rel ? resolvePartPath(this.part, rel.target) : null;
    }

    private anchorRows(anchor: Element): { from: Element | null; to: Element | null } {
        const rowOf = (name: string) => {
            const marker = childElements(anchor).find((c) => localName(c) === name);
            return marker ? childElements(marker).find((c) => localName(c) === 'row') ?? null : null;
        };
        return { from: rowOf('from'), to: rowOf('to') };
    }

    private async anchors(): Promise<{ part: string; doc: Document; anchors: Element[] } | null> {
        const part = await this.drawingPart();
        if (!part) {
            return null;
        }
        const doc = await this.pkg.xml(part);
        const anchors = childElements(doc.documentElement).filter((el) =>
            ['twoCellAnchor', 'oneCellAnchor'].includes(localName(el))
        );
        return { part, doc, anchors };
    }

    private async shiftAnchors(at: number, delta: number) {
        const drawing = await this.anchors();
        if (!drawing) {
            return;
        }
        for (const anchor of drawing.anchors) {
            const { from, to } = this.anchorRows(anchor);
            for (const marker of [from, to]) {
                if (!marker) {
                    continue;
                }
                const row1 = Number(marker.textContent) + 1; // markers are 0-based
                const moved = delta > 0 ? (row1 >= at ? row1 + delta : row1) : row1 >= at - delta ? row1 + delta : row1;
                marker.textContent = String(Math.max(0, moved - 1));
            }
        }
        this.pkg.markDirty(drawing.part);
    }

    /** Removes floating pictures whose top-left corner lies inside the range. */
    async removeAnchorsIn(range: RangeRef) {
        const drawing = await this.anchors();
        if (!drawing) {
            return;
        }
        for (const anchor of drawing.anchors) {
            const { from } = this.anchorRows(anchor);
            const row1 = from ? Number(from.textContent) + 1 : 0;
            if (row1 >= range.top && row1 <= range.bottom) {
                anchor.parentNode?.removeChild(anchor);
                this.pkg.markDirty(drawing.part);
            }
        }
    }

    private async shiftDefinedNames(at: number, delta: number) {
        const wb = await this.pkg.xml('xl/workbook.xml');
        const escaped = this.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/'/g, "''");
        const pattern = new RegExp(
            `('${escaped}'|${/^[A-Za-z_][\w.]*$/.test(this.name) ? escaped : '(?!)'})!(\\$?[A-Za-z]{1,3}\\$?\\d+(?::\\$?[A-Za-z]{1,3}\\$?\\d+)?)`,
            'g'
        );
        let changed = false;
        for (const dn of Array.from(wb.getElementsByTagName('definedName'))) {
            const text = dn.textContent ?? '';
            const next = text.replace(pattern, (_all, sheet: string, ref: string) => {
                changed = true;
                return `${sheet}!${shiftFormula(ref, at, delta)}`;
            });
            if (next !== text) {
                dn.textContent = next;
            }
        }
        if (changed) {
            this.pkg.markDirty('xl/workbook.xml');
        }
    }

    // ---------------------------------------------------------------- pictures

    /**
     * Places a picture over a cell range, scaled to fit and centred, anchored so it moves
     * with the cells. Floating pictures render in every Excel version, unlike pictures
     * placed inside cells, which older desktop builds show as #VALUE!.
     */
    async placePicture(range: RangeRef, data: Buffer, extension: string) {
        const ext = normalizeExtension(extension);
        await this.pkg.ensureContentType({ extension: ext, contentType: IMAGE_TYPES[ext] });

        const media = new Set(this.pkg.listParts('xl/media/'));
        let n = 1;
        while (media.has(`xl/media/image${n}.${ext}`)) {
            n++;
        }
        const mediaPart = `xl/media/image${n}.${ext}`;
        this.pkg.writeBinary(mediaPart, data);

        const drawingPart = await this.ensureDrawingPart();
        const rid = await this.pkg.addRelationship(drawingPart, REL_TYPE.image, `../media/${path.posix.basename(mediaPart)}`);

        // Fit the picture inside the range with a small margin, keeping its aspect ratio.
        const colEdges = this.columnEdges(range.left, range.right + 1);
        const rowEdges = this.rowEdges(range.top, range.bottom + 1);
        const areaW = colEdges[colEdges.length - 1];
        const areaH = rowEdges[rowEdges.length - 1];
        const margin = 4 * EMU_PER_PIXEL;
        const size = imageSize(data) ?? { width: 800, height: 450 };
        const scale = Math.min(
            (areaW - 2 * margin) / (size.width * EMU_PER_PIXEL),
            (areaH - 2 * margin) / (size.height * EMU_PER_PIXEL)
        );
        const cx = Math.max(1, Math.round(size.width * EMU_PER_PIXEL * scale));
        const cy = Math.max(1, Math.round(size.height * EMU_PER_PIXEL * scale));
        const x = Math.round((areaW - cx) / 2);
        const y = Math.round((areaH - cy) / 2);
        const from = locate(colEdges, rowEdges, x, y, range);
        const to = locate(colEdges, rowEdges, x + cx, y + cy, range);

        const doc = await this.pkg.xml(drawingPart);
        const ids = Array.from(doc.getElementsByTagName('*'))
            .filter((el) => localName(el) === 'cNvPr')
            .map((el) => Number(el.getAttribute('id')) || 0);
        const shapeId = Math.max(1, ...ids) + 1;
        const xml =
            `<xdr:twoCellAnchor xmlns:xdr="${NS.drawing}" xmlns:a="${NS.drawingMain}" xmlns:r="${NS.rel}" editAs="oneCell">` +
            `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>${from.colOff}</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>${from.rowOff}</xdr:rowOff></xdr:from>` +
            `<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>${to.colOff}</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>${to.rowOff}</xdr:rowOff></xdr:to>` +
            `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${shapeId}" name="Evidencia ${shapeId}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
            `<xdr:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
            `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>` +
            `<xdr:clientData/></xdr:twoCellAnchor>`;
        const fragment = new DOMParser().parseFromString(xml, 'text/xml').documentElement;
        doc.documentElement.appendChild(doc.importNode(fragment as unknown as Node, true));
        this.pkg.markDirty(drawingPart);
    }

    /** First floating picture whose top-left corner lies inside the range. */
    async pictureIn(range: RangeRef): Promise<CellImage | null> {
        const drawing = await this.anchors();
        if (!drawing) {
            return null;
        }
        const rels = await this.pkg.relationships(drawing.part);
        for (const anchor of drawing.anchors) {
            const { from } = this.anchorRows(anchor);
            const row1 = from ? Number(from.textContent) + 1 : 0;
            if (row1 < range.top || row1 > range.bottom) {
                continue;
            }
            const blip = Array.from(anchor.getElementsByTagName('*')).find((el) => localName(el) === 'blip');
            const rid = blip?.getAttributeNS(NS.rel, 'embed') || blip?.getAttribute('r:embed');
            const rel = rels.find((r) => r.id === rid);
            if (!rel) {
                continue;
            }
            const mediaPart = resolvePartPath(drawing.part, rel.target);
            const data = await this.pkg.binary(mediaPart);
            if (data) {
                return { data, extension: path.posix.extname(mediaPart).slice(1) };
            }
        }
        return null;
    }

    private async ensureDrawingPart(): Promise<string> {
        const existing = await this.drawingPart();
        if (existing) {
            return existing;
        }
        const parts = new Set(this.pkg.listParts('xl/drawings/'));
        let n = 1;
        while (parts.has(`xl/drawings/drawing${n}.xml`)) {
            n++;
        }
        const part = `xl/drawings/drawing${n}.xml`;
        this.pkg.createXml(
            part,
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${NS.drawing}" xmlns:a="${NS.drawingMain}"></xdr:wsDr>`
        );
        await this.pkg.ensureContentType({
            partName: `/${part}`,
            contentType: 'application/vnd.openxmlformats-officedocument.drawing+xml'
        });
        const rid = await this.pkg.addRelationship(this.part, REL_TYPE.drawing, `../drawings/drawing${n}.xml`);
        const el = this.insertWorksheetChild('drawing');
        el.setAttributeNS(NS.rel, 'r:id', rid);
        return part;
    }

    /**
     * Grows a row so wrapped text across columns `fromCol..toCol` is not cut off.
     * Rows are never shrunk, so hand-tuned heights stay.
     */
    fitRowHeight(
        row: number,
        text: string,
        fromCol: number,
        toCol: number,
        pointsPerLine = 15,
        options: { shrink?: boolean; min?: number } = {}
    ) {
        const edges = this.columnEdges(fromCol, toCol + 1);
        // ~7 px per character of 11pt Calibri/Arial.
        const charsPerLine = Math.max(10, edges[edges.length - 1] / EMU_PER_PIXEL / 7);
        const lines = text
            .split('\n')
            .reduce((n, line) => n + Math.max(1, Math.ceil((line.length * 1.08) / charsPerLine)), 0);
        const needed = Math.max(options.min ?? 0, lines * pointsPerLine + 4);
        const rowEl = this.rowEl(row, true)!;
        const fmt = this.doc.getElementsByTagName('sheetFormatPr')[0];
        const current = Number(rowEl.getAttribute('ht')) || Number(fmt?.getAttribute('defaultRowHeight')) || 15;
        if (needed > current || (options.shrink && needed < current)) {
            rowEl.setAttribute('ht', String(Math.min(409, Math.round(needed * 4) / 4)));
            rowEl.setAttribute('customHeight', '1');
            this.touch();
        }
    }

    /** Cumulative EMU offsets of the left edge of each column from `from` up to `to`. */
    private columnEdges(from: number, to: number): number[] {
        const fmt = this.doc.getElementsByTagName('sheetFormatPr')[0];
        const defaultWidth =
            Number(fmt?.getAttribute('defaultColWidth')) || (Number(fmt?.getAttribute('baseColWidth')) || 8) + 0.71;
        const cols = Array.from(this.doc.getElementsByTagName('col'));
        const widthOf = (c: number) => {
            const col = cols.find((el) => Number(el.getAttribute('min')) <= c && c <= Number(el.getAttribute('max')));
            if (col?.getAttribute('hidden') === '1') {
                return 0;
            }
            const chars = col ? Number(col.getAttribute('width')) : defaultWidth;
            return Math.trunc(((256 * chars + Math.trunc(128 / 7)) / 256) * 7) * EMU_PER_PIXEL;
        };
        const edges = [0];
        for (let c = from; c < to; c++) {
            edges.push(edges[edges.length - 1] + widthOf(c));
        }
        return edges;
    }

    /** Cumulative EMU offsets of the top edge of each row from `from` up to `to`. */
    private rowEdges(from: number, to: number): number[] {
        const fmt = this.doc.getElementsByTagName('sheetFormatPr')[0];
        const defaultHeight = Number(fmt?.getAttribute('defaultRowHeight')) || 15;
        const edges = [0];
        for (let r = from; r < to; r++) {
            const row = this.rows.get(r);
            const hidden = row?.getAttribute('hidden') === '1';
            const ht = row?.getAttribute('ht') ? Number(row.getAttribute('ht')) : defaultHeight;
            edges.push(edges[edges.length - 1] + (hidden ? 0 : Math.round(ht * EMU_PER_POINT)));
        }
        return edges;
    }

    /** Reads back the in-cell picture of a cell, if any. */
    async cellImage(ref: string): Promise<CellImage | null> {
        const c = this.cellEl(ref, false);
        const vm = c?.getAttribute('vm');
        if (!vm || !this.pkg.has('xl/metadata.xml')) {
            return null;
        }
        try {
            const meta = await this.pkg.xml('xl/metadata.xml');
            const valueBlocks = childElements(meta.getElementsByTagName('valueMetadata')[0]);
            const rc = valueBlocks[Number(vm) - 1]?.getElementsByTagName('rc')[0];
            const future = Array.from(meta.getElementsByTagName('futureMetadata')).find(
                (f) => f.getAttribute('name') === 'XLRICHVALUE'
            );
            const bk = childElements(future!)[Number(rc?.getAttribute('v'))];
            const rvb = bk ? Array.from(bk.getElementsByTagName('*')).find((e) => localName(e) === 'rvb') : null;
            const rvIndex = Number(rvb?.getAttribute('i'));
            const rels = await this.pkg.relationships('xl/workbook.xml');
            const rvPart = rels.find((r) => r.type === REL_TYPE.richValue);
            const relPart = rels.find((r) => r.type === REL_TYPE.richValueRel);
            if (!rvPart || !relPart) {
                return null;
            }
            const rvDoc = await this.pkg.xml(resolvePartPath('xl/workbook.xml', rvPart.target));
            const rv = rvDoc.getElementsByTagName('rv')[rvIndex];
            const relIndex = Number(rv?.getElementsByTagName('v')[0]?.textContent);
            const relXmlPart = resolvePartPath('xl/workbook.xml', relPart.target);
            const relDoc = await this.pkg.xml(relXmlPart);
            const relEl = relDoc.getElementsByTagName('rel')[relIndex];
            const rid = relEl?.getAttributeNS(NS.rel, 'id') || relEl?.getAttribute('r:id');
            const imageRel = (await this.pkg.relationships(relXmlPart)).find((r) => r.id === rid);
            if (!imageRel) {
                return null;
            }
            const mediaPart = resolvePartPath(relXmlPart, imageRel.target);
            const data = await this.pkg.binary(mediaPart);
            return data ? { data, extension: path.posix.extname(mediaPart).slice(1) } : null;
        } catch {
            return null;
        }
    }
}

// -------------------------------------------------------------------- picture helpers

const EMU_PER_PIXEL = 9525;
const EMU_PER_POINT = 12700;

const IMAGE_TYPES: Record<string, string> = {
    png: 'image/png',
    jpeg: 'image/jpeg',
    gif: 'image/gif'
};

function normalizeExtension(extension: string): string {
    const ext = extension.toLowerCase().replace(/^\./, '');
    if (ext === 'jpg' || ext === 'jpeg') {
        return 'jpeg';
    }
    return ext === 'gif' ? 'gif' : 'png';
}

/** Converts an EMU offset inside the range into a 0-based cell marker. */
function locate(colEdges: number[], rowEdges: number[], x: number, y: number, range: RangeRef) {
    const pick = (edges: number[], v: number) => {
        let i = 0;
        while (i < edges.length - 2 && edges[i + 1] <= v) {
            i++;
        }
        return { index: i, offset: Math.max(0, v - edges[i]) };
    };
    const c = pick(colEdges, x);
    const r = pick(rowEdges, y);
    return { col: range.left - 1 + c.index, colOff: c.offset, row: range.top - 1 + r.index, rowOff: r.offset };
}

/** Pixel size of a PNG, JPEG or GIF image. */
export function imageSize(data: Buffer): { width: number; height: number } | null {
    if (data.length > 24 && data.readUInt32BE(0) === 0x89504e47) {
        return { width: data.readUInt32BE(16), height: data.readUInt32BE(20) };
    }
    if (data.length > 10 && data.toString('ascii', 0, 3) === 'GIF') {
        return { width: data.readUInt16LE(6), height: data.readUInt16LE(8) };
    }
    if (data.length > 4 && data[0] === 0xff && data[1] === 0xd8) {
        let i = 2;
        while (i + 9 < data.length) {
            if (data[i] !== 0xff) {
                i++;
                continue;
            }
            const marker = data[i + 1];
            const length = data.readUInt16BE(i + 2);
            if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) {
                return { width: data.readUInt16BE(i + 7), height: data.readUInt16BE(i + 5) };
            }
            i += 2 + length;
        }
    }
    return null;
}
