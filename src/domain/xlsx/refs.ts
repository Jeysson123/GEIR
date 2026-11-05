// A1-style reference helpers shared by the worksheet editor.

export interface CellRef {
    col: number; // 1-based
    row: number; // 1-based
}

export interface RangeRef {
    top: number;
    left: number;
    bottom: number;
    right: number;
}

export function colToNumber(col: string): number {
    let n = 0;
    for (const ch of col.toUpperCase()) {
        n = n * 26 + (ch.charCodeAt(0) - 64);
    }
    return n;
}

export function numberToCol(n: number): string {
    let s = '';
    while (n > 0) {
        const m = (n - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}

export function parseCell(ref: string): CellRef {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref.trim());
    if (!m) {
        throw new Error(`Invalid cell reference: ${ref}`);
    }
    return { col: colToNumber(m[1]), row: Number(m[2]) };
}

export function cellName(col: number, row: number): string {
    return `${numberToCol(col)}${row}`;
}

export function parseRange(ref: string): RangeRef {
    const [a, b] = ref.split(':');
    const start = parseCell(a);
    const end = b ? parseCell(b) : start;
    return {
        top: Math.min(start.row, end.row),
        bottom: Math.max(start.row, end.row),
        left: Math.min(start.col, end.col),
        right: Math.max(start.col, end.col)
    };
}

export function formatRange(r: RangeRef): string {
    const a = cellName(r.left, r.top);
    const b = cellName(r.right, r.bottom);
    return a === b ? a : `${a}:${b}`;
}

/**
 * Moves a range after rows were inserted (delta > 0) or deleted (delta < 0) at `at`.
 * Returns null when the whole range falls inside deleted rows.
 */
export function shiftRange(r: RangeRef, at: number, delta: number): RangeRef | null {
    if (delta === 0) {
        return { ...r };
    }

    if (delta > 0) {
        const top = r.top >= at ? r.top + delta : r.top;
        const bottom = r.bottom >= at ? r.bottom + delta : r.bottom;
        return { ...r, top, bottom };
    }

    const delStart = at;
    const delEnd = at - delta - 1;
    const mapRow = (row: number) => (row > delEnd ? row + delta : row);

    if (r.top >= delStart && r.bottom <= delEnd) {
        return null;
    }

    let top = r.top;
    let bottom = r.bottom;
    if (top >= delStart && top <= delEnd) {
        top = delEnd + 1;
    }
    if (bottom >= delStart && bottom <= delEnd) {
        bottom = delStart - 1;
    }
    return { ...r, top: mapRow(top), bottom: mapRow(bottom) };
}

/** Shifts every range in a space-separated sqref list, dropping deleted ones. */
export function shiftSqref(sqref: string, at: number, delta: number): string {
    return sqref
        .split(/\s+/)
        .filter(Boolean)
        .map((part) => {
            const shifted = shiftRange(parseRange(part), at, delta);
            return shifted ? formatRange(shifted) : '';
        })
        .filter(Boolean)
        .join(' ');
}

/**
 * Shifts same-sheet A1 references inside a formula. References qualified with another
 * sheet name are left alone; string literals are skipped.
 */
export function shiftFormula(formula: string, at: number, delta: number): string {
    let out = '';
    let i = 0;
    while (i < formula.length) {
        const ch = formula[i];
        if (ch === '"') {
            const end = formula.indexOf('"', i + 1);
            const stop = end === -1 ? formula.length : end + 1;
            out += formula.slice(i, stop);
            i = stop;
            continue;
        }
        if (ch === "'") {
            // Quoted sheet name: copy it and the reference that follows untouched.
            const end = formula.indexOf("'!", i + 1);
            if (end !== -1) {
                const m = /^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?/.exec(formula.slice(end + 2));
                const stop = end + 2 + (m ? m[0].length : 0);
                out += formula.slice(i, stop);
                i = stop;
                continue;
            }
        }
        const rest = formula.slice(i);
        const sheetQualified = /^[A-Za-z_][\w.]*!\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?/.exec(rest);
        if (sheetQualified && (i === 0 || !/[\w.]/.test(formula[i - 1]))) {
            out += sheetQualified[0];
            i += sheetQualified[0].length;
            continue;
        }
        const ref = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w(])/.exec(rest);
        if (ref && (i === 0 || !/[\w.$]/.test(formula[i - 1]))) {
            const row = Number(ref[4]);
            const newRow = delta > 0 ? (row >= at ? row + delta : row) : row > at - delta - 1 ? row + delta : row;
            out += `${ref[1]}${ref[2]}${ref[3]}${newRow}`;
            i += ref[0].length;
            continue;
        }
        out += ch;
        i++;
    }
    return out;
}
