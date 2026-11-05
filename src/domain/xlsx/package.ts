import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import * as path from 'path';
import { cellName, colToNumber } from './refs';
import { Worksheet } from './worksheet';

export const NS = {
    main: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    pkgRel: 'http://schemas.openxmlformats.org/package/2006/relationships',
    contentTypes: 'http://schemas.openxmlformats.org/package/2006/content-types',
    drawing: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    drawingMain: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    richData: 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata',
    richValueRel: 'http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel'
};

export const REL_TYPE = {
    hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
    image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    drawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
    calcChain: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain',
    sheetMetadata: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata',
    richValue: 'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue',
    richValueStructure: 'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure',
    richValueTypes: 'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes',
    richValueRel: 'http://schemas.microsoft.com/office/2022/10/relationships/richValueRel'
};

export interface Relationship {
    id: string;
    type: string;
    target: string;
    external: boolean;
}

/** Joins a relationship target with the folder of the part that owns it. */
export function resolvePartPath(ownerPart: string, target: string): string {
    if (target.startsWith('/')) {
        return target.slice(1);
    }
    return path.posix.normalize(path.posix.join(path.posix.dirname(ownerPart), target));
}

function relsPathFor(part: string): string {
    return path.posix.join(path.posix.dirname(part), '_rels', `${path.posix.basename(part)}.rels`);
}

/**
 * Edits an .xlsx package at the XML level. Parts that are never touched are written
 * back byte for byte, so features other libraries drop (in-cell pictures, custom XML,
 * protected ranges, x14 extensions) survive the round trip.
 */
export class XlsxPackage {
    private docs = new Map<string, Document>();
    private dirty = new Set<string>();
    private sheets = new Map<string, Worksheet>();
    private sharedStringsCache: string[] | null = null;

    private constructor(private zip: JSZip) {}

    static async load(data: Buffer | Uint8Array): Promise<XlsxPackage> {
        const zip = await JSZip.loadAsync(data);
        return new XlsxPackage(zip);
    }

    has(part: string): boolean {
        return this.zip.file(part) !== null || this.docs.has(part);
    }

    async xml(part: string): Promise<Document> {
        const cached = this.docs.get(part);
        if (cached) {
            return cached;
        }
        const file = this.zip.file(part);
        if (!file) {
            throw new Error(`Part not found in workbook: ${part}`);
        }
        const text = await file.async('string');
        const doc = new DOMParser().parseFromString(text, 'text/xml') as unknown as Document;
        this.docs.set(part, doc);
        return doc;
    }

    createXml(part: string, xml: string): Document {
        const doc = new DOMParser().parseFromString(xml, 'text/xml') as unknown as Document;
        this.docs.set(part, doc);
        this.dirty.add(part);
        return doc;
    }

    markDirty(part: string) {
        this.dirty.add(part);
    }

    async binary(part: string): Promise<Buffer | null> {
        const file = this.zip.file(part);
        return file ? file.async('nodebuffer') : null;
    }

    writeBinary(part: string, data: Buffer) {
        this.zip.file(part, data);
    }

    removePart(part: string) {
        this.zip.remove(part);
        this.docs.delete(part);
        this.dirty.delete(part);
    }

    listParts(prefix: string): string[] {
        const names: string[] = [];
        this.zip.forEach((relativePath) => {
            if (relativePath.startsWith(prefix)) {
                names.push(relativePath);
            }
        });
        return names;
    }

    async relationships(part: string): Promise<Relationship[]> {
        const relsPath = relsPathFor(part);
        if (!this.has(relsPath)) {
            return [];
        }
        const doc = await this.xml(relsPath);
        return Array.from(doc.getElementsByTagName('Relationship')).map((el) => ({
            id: el.getAttribute('Id') ?? '',
            type: el.getAttribute('Type') ?? '',
            target: el.getAttribute('Target') ?? '',
            external: el.getAttribute('TargetMode') === 'External'
        }));
    }

    async addRelationship(part: string, type: string, target: string, external = false): Promise<string> {
        const relsPath = relsPathFor(part);
        if (!this.has(relsPath)) {
            this.createXml(
                relsPath,
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS.pkgRel}"></Relationships>`
            );
        }
        const doc = await this.xml(relsPath);
        const root = doc.documentElement;
        const used = new Set(Array.from(root.getElementsByTagName('Relationship')).map((el) => el.getAttribute('Id')));
        let n = used.size + 1;
        while (used.has(`rId${n}`)) {
            n++;
        }
        const id = `rId${n}`;
        const el = doc.createElementNS(NS.pkgRel, 'Relationship');
        el.setAttribute('Id', id);
        el.setAttribute('Type', type);
        el.setAttribute('Target', target);
        if (external) {
            el.setAttribute('TargetMode', 'External');
        }
        root.appendChild(el);
        this.markDirty(relsPath);
        return id;
    }

    async removeRelationship(part: string, id: string) {
        const relsPath = relsPathFor(part);
        if (!this.has(relsPath)) {
            return;
        }
        const doc = await this.xml(relsPath);
        for (const el of Array.from(doc.getElementsByTagName('Relationship'))) {
            if (el.getAttribute('Id') === id) {
                el.parentNode?.removeChild(el);
                this.markDirty(relsPath);
            }
        }
    }

    async ensureContentType(options: { extension?: string; partName?: string; contentType: string }) {
        const doc = await this.xml('[Content_Types].xml');
        const root = doc.documentElement;
        if (options.extension) {
            const ext = options.extension.toLowerCase();
            const exists = Array.from(root.getElementsByTagName('Default')).some(
                (el) => (el.getAttribute('Extension') ?? '').toLowerCase() === ext
            );
            if (!exists) {
                const el = doc.createElementNS(NS.contentTypes, 'Default');
                el.setAttribute('Extension', ext);
                el.setAttribute('ContentType', options.contentType);
                root.insertBefore(el, root.firstChild);
                this.markDirty('[Content_Types].xml');
            }
        }
        if (options.partName) {
            const exists = Array.from(root.getElementsByTagName('Override')).some(
                (el) => el.getAttribute('PartName') === options.partName
            );
            if (!exists) {
                const el = doc.createElementNS(NS.contentTypes, 'Override');
                el.setAttribute('PartName', options.partName);
                el.setAttribute('ContentType', options.contentType);
                root.appendChild(el);
                this.markDirty('[Content_Types].xml');
            }
        }
    }

    async sharedStrings(): Promise<string[]> {
        if (this.sharedStringsCache) {
            return this.sharedStringsCache;
        }
        const part = 'xl/sharedStrings.xml';
        if (!this.has(part)) {
            this.sharedStringsCache = [];
            return this.sharedStringsCache;
        }
        const doc = await this.xml(part);
        this.sharedStringsCache = Array.from(doc.getElementsByTagName('si')).map((si) =>
            Array.from(si.getElementsByTagName('t'))
                // Phonetic runs (<rPh>) are not part of the visible text.
                .filter((t) => (t.parentNode as Element | null)?.nodeName !== 'rPh')
                .map((t) => t.textContent ?? '')
                .join('')
        );
        return this.sharedStringsCache;
    }

    async sheetNames(): Promise<string[]> {
        const wb = await this.xml('xl/workbook.xml');
        return Array.from(wb.getElementsByTagName('sheet')).map((el) => el.getAttribute('name') ?? '');
    }

    /** Finds a worksheet by name, ignoring case, accents and surrounding spaces. */
    async sheet(name: string): Promise<Worksheet> {
        const wb = await this.xml('xl/workbook.xml');
        const wanted = normalizeName(name);
        const el = Array.from(wb.getElementsByTagName('sheet')).find(
            (s) => normalizeName(s.getAttribute('name') ?? '') === wanted
        );
        if (!el) {
            throw new Error(`Sheet "${name}" not found in workbook`);
        }
        const rid = el.getAttributeNS(NS.rel, 'id') ?? el.getAttribute('r:id') ?? '';
        const rel = (await this.relationships('xl/workbook.xml')).find((r) => r.id === rid);
        if (!rel) {
            throw new Error(`Relationship ${rid} for sheet "${name}" not found`);
        }
        const part = resolvePartPath('xl/workbook.xml', rel.target);
        let sheet = this.sheets.get(part);
        if (!sheet) {
            const index = Array.from(wb.getElementsByTagName('sheet')).indexOf(el);
            sheet = new Worksheet(this, part, el.getAttribute('name') ?? name, index);
            await sheet.init();
            this.sheets.set(part, sheet);
        }
        return sheet;
    }

    /**
     * Drops the calculation chain so Excel rebuilds it on open. Required after moving
     * formula cells, harmless otherwise.
     */
    async dropCalcChain() {
        const rels = await this.relationships('xl/workbook.xml');
        const rel = rels.find((r) => r.type === REL_TYPE.calcChain);
        if (!rel) {
            return;
        }
        const part = resolvePartPath('xl/workbook.xml', rel.target);
        this.removePart(part);
        await this.removeRelationship('xl/workbook.xml', rel.id);
        const ct = await this.xml('[Content_Types].xml');
        for (const el of Array.from(ct.getElementsByTagName('Override'))) {
            if (el.getAttribute('PartName') === `/${part}`) {
                el.parentNode?.removeChild(el);
                this.markDirty('[Content_Types].xml');
            }
        }
    }

    /** Asks Excel to recalculate every formula when the file opens (e.g. totals of edited cells). */
    async recalculateOnOpen() {
        const wb = await this.xml('xl/workbook.xml');
        let calcPr = wb.getElementsByTagName('calcPr')[0];
        if (!calcPr) {
            calcPr = wb.createElementNS(NS.main, 'calcPr');
            // calcPr goes after sheets/definedNames; before any extension list.
            const ext = wb.getElementsByTagName('extLst')[0];
            wb.documentElement.insertBefore(calcPr, ext && ext.parentNode === wb.documentElement ? ext : null);
        }
        calcPr.setAttribute('fullCalcOnLoad', '1');
        this.markDirty('xl/workbook.xml');
    }

    /**
     * Leaves only the named worksheet (the derived files: one GEIR sheet per workbook).
     * List validations and formulas that point to removed sheets become fixed values,
     * so Excel opens the result without repairs.
     */
    async keepOnlySheet(name: string) {
        const wbPart = 'xl/workbook.xml';
        const wb = await this.xml(wbPart);
        const rels = await this.relationships(wbPart);
        const sheetEls = Array.from(wb.getElementsByTagName('sheet'));
        const keep = sheetEls.find((s) => normalizeName(s.getAttribute('name') ?? '') === normalizeName(name));
        if (!keep) {
            throw new Error(`Sheet "${name}" not found in workbook`);
        }
        const keepIndex = sheetEls.indexOf(keep);
        const partOf = (el: Element) => {
            const rid = el.getAttributeNS(NS.rel, 'id') ?? el.getAttribute('r:id') ?? '';
            const rel = rels.find((r) => r.id === rid);
            return { rid, part: rel ? resolvePartPath(wbPart, rel.target) : '' };
        };
        const keepPart = partOf(keep).part;
        const removed = sheetEls.filter((el) => el !== keep);
        const removedNames = removed.map((el) => el.getAttribute('name') ?? '');

        // Values of list validations / formulas that point to removed sheets, read first.
        const keepDoc = await this.xml(keepPart);
        const listValues = async (ref: string): Promise<string[] | null> => {
            const m = /^'?(.+?)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(ref.trim());
            if (!m || !removedNames.some((n) => normalizeName(n) === normalizeName(m[1]))) {
                return null;
            }
            const source = await this.sheet(m[1]);
            const values: string[] = [];
            const col1 = colToNumber(m[2]);
            const col2 = colToNumber(m[4] ?? m[2]);
            for (let r = Number(m[3]); r <= Number(m[5] ?? m[3]); r++) {
                for (let c = col1; c <= col2; c++) {
                    const v = await source.displayText(cellName(c, r));
                    if (v) {
                        values.push(v.replace(/"/g, "'").replace(/,/g, ' '));
                    }
                }
            }
            return values;
        };
        for (const f of [
            ...Array.from(keepDoc.getElementsByTagName('formula1')),
            ...Array.from(keepDoc.getElementsByTagName('xm:f'))
        ]) {
            const values = await listValues(f.textContent ?? '');
            if (values) {
                f.textContent = `"${values.join(',')}"`;
            }
        }
        const pointsToRemoved = (text: string) =>
            removedNames.some((n) => text.includes(`${n}!`) || text.includes(`'${n.replace(/'/g, "''")}'!`));
        for (const f of Array.from(keepDoc.getElementsByTagName('f'))) {
            if (pointsToRemoved(f.textContent ?? '')) {
                f.parentNode?.removeChild(f); // keeps the cached <v>
            }
        }
        this.markDirty(keepPart);

        // Workbook: sheets, defined names, active tab.
        for (const el of removed) {
            const { rid, part } = partOf(el);
            el.parentNode?.removeChild(el);
            await this.removeRelationship(wbPart, rid);
            if (part && part !== keepPart) {
                for (const rel of await this.relationships(part)) {
                    if (rel.type === REL_TYPE.drawing) {
                        const drawing = resolvePartPath(part, rel.target);
                        this.removePart(drawing);
                        this.removePart(relsPathFor(drawing));
                        await this.removeOverride(`/${drawing}`);
                    }
                }
                this.removePart(part);
                this.removePart(relsPathFor(part));
                await this.removeOverride(`/${part}`);
            }
        }
        keep.removeAttribute('state');
        for (const dn of Array.from(wb.getElementsByTagName('definedName'))) {
            const local = dn.getAttribute('localSheetId');
            if ((local !== null && Number(local) !== keepIndex) || pointsToRemoved(dn.textContent ?? '')) {
                dn.parentNode?.removeChild(dn);
            } else if (local !== null) {
                dn.setAttribute('localSheetId', '0');
            }
        }
        const names = wb.getElementsByTagName('definedNames')[0];
        if (names && !names.getElementsByTagName('definedName').length) {
            names.parentNode?.removeChild(names);
        }
        for (const view of Array.from(wb.getElementsByTagName('workbookView'))) {
            view.setAttribute('activeTab', '0');
            view.removeAttribute('firstSheet');
        }
        for (const view of Array.from(keepDoc.getElementsByTagName('sheetView'))) {
            view.setAttribute('tabSelected', '1');
        }
        this.markDirty(wbPart);

        // docProps/app.xml lists the sheet titles; it is optional, so drop it.
        const rootRels = await this.relationships('');
        for (const rel of rootRels.filter((r) => r.target.replace(/^\//, '') === 'docProps/app.xml')) {
            await this.removeRelationship('', rel.id);
        }
        this.removePart('docProps/app.xml');
        await this.removeOverride('/docProps/app.xml');
        await this.dropCalcChain();

        // In-cell pictures (rich data) belong to the workbook: drop them when the kept sheet
        // has none, then every picture no remaining part points to.
        const usesRichData = Array.from(keepDoc.getElementsByTagName('c')).some((c) => c.hasAttribute('vm'));
        if (!usesRichData) {
            const richTypes = [
                REL_TYPE.richValue,
                REL_TYPE.richValueStructure,
                REL_TYPE.richValueTypes,
                REL_TYPE.richValueRel,
                REL_TYPE.sheetMetadata
            ];
            for (const rel of (await this.relationships(wbPart)).filter((r) => richTypes.includes(r.type))) {
                const part = resolvePartPath(wbPart, rel.target);
                this.removePart(part);
                this.removePart(relsPathFor(part));
                await this.removeOverride(`/${part}`);
                await this.removeRelationship(wbPart, rel.id);
            }
        }
        await this.pruneMedia();
        this.sheets.clear();
    }

    /** Removes .rels of parts that no longer exist and pictures nothing points to. */
    private async pruneMedia() {
        const used = new Set<string>();
        for (const rels of this.listParts('').filter((p) => p.endsWith('.rels'))) {
            // "_rels/.rels" belongs to the package itself; "x/_rels/y.xml.rels" to "x/y.xml".
            const ownerPart =
                rels === '_rels/.rels'
                    ? ''
                    : path.posix.join(path.posix.dirname(path.posix.dirname(rels)), path.posix.basename(rels).replace(/\.rels$/, ''));
            if (ownerPart && !this.has(ownerPart)) {
                this.removePart(rels);
                continue;
            }
            for (const rel of await this.relationships(ownerPart)) {
                if (!rel.external) {
                    used.add(resolvePartPath(ownerPart, rel.target));
                }
            }
        }
        for (const media of this.listParts('xl/media/')) {
            if (!used.has(media)) {
                this.removePart(media);
            }
        }
    }

    private async removeOverride(partName: string) {
        const ct = await this.xml('[Content_Types].xml');
        for (const el of Array.from(ct.getElementsByTagName('Override'))) {
            if (el.getAttribute('PartName') === partName) {
                el.parentNode?.removeChild(el);
                this.markDirty('[Content_Types].xml');
            }
        }
    }

    async save(): Promise<Buffer> {
        const serializer = new XMLSerializer();
        for (const part of this.dirty) {
            const doc = this.docs.get(part);
            if (!doc) {
                continue;
            }
            let text = serializer.serializeToString(doc as any);
            if (!text.startsWith('<?xml')) {
                text = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${text}`;
            }
            this.zip.file(part, text);
        }
        this.dirty.clear();
        return this.zip.generateAsync({
            type: 'nodebuffer',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 },
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
    }
}

export function normalizeName(name: string): string {
    return name
        .normalize('NFD')
        .replace(/[̀-ͯ]/g, '')
        .trim()
        .toLowerCase();
}
