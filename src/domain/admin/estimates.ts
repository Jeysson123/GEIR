import { PlanRow } from '../document';
import { PlanPhase } from './config';

/**
 * "Plan de Trabajo" from two dates: the first user story of the GEIR and the last one
 * added. Phases run back to back on working days (Monday–Friday, minus holidays):
 * fixed phases take their configured days; the development phase takes the working days
 * left until the last story's date (at least `minimo`).
 */
export function computePlan(
    phases: PlanPhase[],
    firstDate: string,
    lastDate: string,
    holidays: string[] = []
): { rows: PlanRow[]; totalDays: number; release: string } {
    const off = new Set(holidays);
    const isWorkday = (d: Date) => d.getUTCDay() !== 0 && d.getUTCDay() !== 6 && !off.has(iso(d));
    const nextWorkday = (d: Date) => {
        const n = new Date(d);
        while (!isWorkday(n)) {
            n.setUTCDate(n.getUTCDate() + 1);
        }
        return n;
    };
    const afterDays = (start: Date, days: number) => {
        // Last working day when `days` working days start on `start`.
        const n = new Date(start);
        for (let left = days - 1; left > 0; ) {
            n.setUTCDate(n.getUTCDate() + 1);
            if (isWorkday(n)) {
                left--;
            }
        }
        return n;
    };
    const workdaysBetween = (from: Date, to: Date) => {
        let count = 0;
        for (const n = new Date(from); n <= to; n.setUTCDate(n.getUTCDate() + 1)) {
            if (isWorkday(n)) {
                count++;
            }
        }
        return count;
    };

    const last = parse(lastDate) ?? parse(firstDate) ?? today();
    let cursor = nextWorkday(parse(firstDate) ?? today());
    const rows: PlanRow[] = [];
    let release = '';
    for (const phase of phases) {
        let days = phase.dias ?? 0;
        if (phase.desarrollo) {
            days = Math.max(phase.minimo ?? 1, workdaysBetween(cursor, last));
        }
        if (days <= 0) {
            rows.push({ activity: phase.actividad, days: 0, start: '', end: '' });
            continue;
        }
        const end = afterDays(cursor, days);
        rows.push({ activity: phase.actividad, days, start: iso(cursor), end: iso(end) });
        release = iso(end);
        const next = new Date(end);
        next.setUTCDate(next.getUTCDate() + 1);
        cursor = nextWorkday(next);
    }
    return { rows, totalDays: rows.reduce((sum, r) => sum + r.days, 0), release };
}

function parse(value: string | undefined): Date | null {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(value ?? '');
    return m ? new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]))) : null;
}

function today(): Date {
    const now = new Date();
    return new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
}

function iso(d: Date): string {
    return d.toISOString().slice(0, 10);
}

/** First non-empty date among the given work item fields, in order. */
export function pickDate(item: Record<string, any> | null | undefined, fields: string[]): string {
    for (const field of fields) {
        const value = item?.[field];
        if (value) {
            return String(value).slice(0, 10);
        }
    }
    return '';
}

const MONTHS = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const DAYS = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];

/** 2026-10-06 → "06/10/2026" (approval e-mails). */
export function shortDate(isoDate: string): string {
    const d = parse(isoDate);
    return d ? `${pad(d.getUTCDate())}/${pad(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}` : isoDate;
}

/** 2026-10-06 → "Martes 06 de octubre del 2026" (FORO e-mail). */
export function longDate(isoDate: string): string {
    const d = parse(isoDate);
    return d ? `${DAYS[d.getUTCDay()]} ${pad(d.getUTCDate())} de ${MONTHS[d.getUTCMonth()]} del ${d.getUTCFullYear()}` : isoDate;
}

function pad(n: number): string {
    return String(n).padStart(2, '0');
}
