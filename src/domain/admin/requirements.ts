/** One requirement ↔ user story pair given by the admin, e.g. 2026-REQ-0101 ↔ 100002. */
export interface RequirementPair {
    requirement: string;
    storyId: string;
    title: string; // optional description typed with the pair
}

const REQUIREMENT = /\b\d{4}-[A-Za-z]+-\d+(?:-\d+)?\b/;
const STORY_ID = /\b\d{4,8}\b/;

/**
 * Reads the requirement list the admin pastes. Accepts, one pair per line:
 *   2026-REQ-0101 100002
 *   2026-REQ-0101: 100002 - Nuevo reporte
 *   rows copied from the sprint Excel (columns separated by tabs)
 * or a JSON array: [["2026-REQ-0101","100002"]] / [{"requerimiento":"…","id":"…"}].
 * The pairing is exactly the one given: nothing is guessed.
 */
export function parseRequirementPairs(text: string): { pairs: RequirementPair[]; errors: string[] } {
    const pairs: RequirementPair[] = [];
    const errors: string[] = [];
    const add = (requirement: string, storyId: string, title: string, source: string) => {
        requirement = requirement.trim().toUpperCase();
        storyId = storyId.trim();
        if (!REQUIREMENT.test(requirement) || !/^\d{4,8}$/.test(storyId)) {
            errors.push(`No se entiende: "${source.trim().slice(0, 80)}" (se espera requerimiento e ID de historia).`);
            return;
        }
        const existing = pairs.find((p) => p.requirement === requirement);
        if (existing && existing.storyId !== storyId) {
            errors.push(`${requirement} aparece con dos historias (${existing.storyId} y ${storyId}).`);
            return;
        }
        const sameStory = pairs.find((p) => p.storyId === storyId && p.requirement !== requirement);
        if (sameStory) {
            errors.push(`La historia ${storyId} aparece en dos requerimientos (${sameStory.requirement} y ${requirement}).`);
            return;
        }
        if (!existing) {
            pairs.push({ requirement, storyId, title: title.trim() });
        }
    };

    const trimmed = text.trim();
    if (!trimmed) {
        return { pairs, errors };
    }
    if (trimmed.startsWith('[')) {
        try {
            const items = JSON.parse(trimmed);
            for (const item of Array.isArray(items) ? items : []) {
                if (Array.isArray(item)) {
                    add(String(item[0] ?? ''), String(item[1] ?? ''), String(item[2] ?? ''), JSON.stringify(item));
                } else if (item && typeof item === 'object') {
                    const req = item.requerimiento ?? item.requirement ?? '';
                    const id = item.id ?? item.historia ?? item.storyId ?? item.azure ?? '';
                    add(String(req), String(id), String(item.titulo ?? item.title ?? item.descripcion ?? ''), JSON.stringify(item));
                }
            }
            return { pairs, errors };
        } catch {
            errors.push('El texto empieza con "[" pero no es un arreglo JSON válido; se lee línea por línea.');
        }
    }

    for (const line of trimmed.split(/\r?\n/)) {
        if (!line.trim()) {
            continue;
        }
        const req = REQUIREMENT.exec(line);
        if (!req) {
            // Header rows of a pasted table ("Requerimiento  ID Azure DevOps …") are skipped.
            if (/\d{4,8}/.test(line)) {
                errors.push(`Falta el requerimiento en: "${line.trim().slice(0, 80)}".`);
            }
            continue;
        }
        const rest = line.slice(0, req.index) + ' '.repeat(req[0].length) + line.slice(req.index + req[0].length);
        // The story ID is the first 4–8 digit number that is not the row counter.
        const cells = line.split('\t');
        let id = '';
        let title = '';
        if (cells.length >= 3) {
            const reqCell = cells.findIndex((c) => REQUIREMENT.test(c));
            const idCell = cells.findIndex((c, i) => i > reqCell && /^\s*\d{4,8}\s*$/.test(c));
            id = idCell >= 0 ? cells[idCell].trim() : '';
            title = idCell >= 0 ? (cells[idCell + 1] ?? '').trim() : '';
        }
        if (!id) {
            // Usually after the requirement ("REQ: ID - título"); else anywhere ("ID REQ").
            const after = rest.slice(req.index);
            const m = STORY_ID.exec(after);
            if (m) {
                id = m[0];
                title = after.slice(m.index + m[0].length).replace(/^\s*[:\-–]\s*/, '');
            } else {
                id = STORY_ID.exec(rest)?.[0] ?? '';
            }
        }
        add(req[0], id, title, line);
    }
    return { pairs, errors };
}
