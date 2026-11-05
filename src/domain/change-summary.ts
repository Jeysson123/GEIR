import * as path from 'path';
import type { FileChange } from '../infrastructure/gitlab/gitlab.client';

export type Layer = 'backend' | 'frontend' | 'base de datos' | 'configuración';

export interface FileSummary {
    path: string;
    name: string;
    layer: Layer;
    status: FileChange['status'];
    additions: number;
    deletions: number;
    added: string[];
    removed: string[];
    modified: string[];
    endpoints: string[];
    sentence: string;
}

const BACKEND = new Set(['.java', '.kt', '.groovy', '.scala', '.py', '.cs', '.go', '.rb', '.php']);
const FRONTEND = new Set(['.js', '.jsx', '.ts', '.tsx', '.vue', '.css', '.scss', '.less', '.html', '.jsp']);
const DATABASE = new Set(['.sql', '.ddl', '.plsql', '.pks', '.pkb']);

const KEYWORDS = new Set([
    'if', 'for', 'while', 'switch', 'catch', 'return', 'new', 'else', 'try', 'do', 'synchronized',
    'super', 'this', 'function', 'typeof', 'await', 'throw', 'case', 'constructor', 'render'
]);

export function layerOf(file: string): Layer {
    const ext = path.extname(file).toLowerCase();
    if (BACKEND.has(ext)) return 'backend';
    if (FRONTEND.has(ext)) return 'frontend';
    if (DATABASE.has(ext)) return 'base de datos';
    return 'configuración';
}

/** Names declared on a line: Java/C# methods & classes, JS functions, components, methods. */
function declarations(line: string, layer: Layer): string[] {
    const out: string[] = [];
    const push = (name: string | undefined, suffix = '()') => {
        if (name && !KEYWORDS.has(name)) out.push(`${name}${suffix}`);
    };
    const type = /\b(class|interface|enum|record)\s+([A-Z]\w*)/.exec(line);
    if (type) {
        out.push(type[2]);
        return out;
    }
    if (layer === 'backend') {
        const method = /^\s*(?:@\w+\s+)*(?:(?:public|private|protected|static|final|abstract|synchronized|default|override|async|internal|virtual)\s+)+[\w<>\[\],.?\s]*?\s(\w+)\s*\([^;]*$/.exec(line);
        push(method?.[1]);
        return out;
    }
    if (layer === 'frontend') {
        push(/\bfunction\s*\*?\s*(\w+)\s*\(/.exec(line)?.[1]);
        push(/\b(?:const|let|var)\s+(\w+)\s*=\s*(?:async\s*)?(?:\([^)]*\)|\w+)\s*=>/.exec(line)?.[1]);
        push(/^\s*(?:static\s+)?(?:async\s+)?(\w+)\s*\([^)]*\)\s*\{\s*$/.exec(line)?.[1]);
        push(/^\s*(\w+)\s*=\s*(?:async\s*)?\([^)]*\)\s*=>/.exec(line)?.[1]);
    }
    return out;
}

/** The method/function named in a hunk header ("@@ -1,2 +1,3 @@ public void foo(").*/
function hunkContext(header: string): string | null {
    const ctx = header.replace(/^@@[^@]*@@\s*/, '');
    if (!ctx) return null;
    const m = /(\w+)\s*\(/.exec(ctx.replace(/\b(if|for|while|switch|catch|return|new)\s*\(/g, ''));
    return m && !KEYWORDS.has(m[1]) ? `${m[1]}()` : null;
}

function unique(items: string[]): string[] {
    return [...new Set(items)];
}

function list(items: string[], max = 4): string {
    const shown = items.slice(0, max);
    const rest = items.length - shown.length;
    const text = shown.length > 1 ? `${shown.slice(0, -1).join(', ')} y ${shown[shown.length - 1]}` : shown[0] ?? '';
    return rest > 0 ? `${text} (y ${rest} más)` : text;
}

export function summarizeFile(change: FileChange): FileSummary {
    const layer = layerOf(change.path);
    const name = path.posix.basename(change.path);
    const added: string[] = [];
    const removed: string[] = [];
    const touched: string[] = [];
    const endpoints: string[] = [];
    let additions = 0;
    let deletions = 0;

    for (const line of change.diff.split('\n')) {
        if (line.startsWith('@@')) {
            const ctx = hunkContext(line);
            if (ctx) touched.push(ctx);
        } else if (line.startsWith('+') && !line.startsWith('+++')) {
            additions++;
            added.push(...declarations(line.slice(1), layer));
            const re = /["'`](\/[\w\-{}.$]+(?:\/[\w\-{}.$]+)+)["'`]/g;
            let m: RegExpExecArray | null;
            while ((m = re.exec(line))) endpoints.push(m[1]);
        } else if (line.startsWith('-') && !line.startsWith('---')) {
            deletions++;
            removed.push(...declarations(line.slice(1), layer));
        }
    }

    const addedSet = unique(added);
    const removedSet = unique(removed);
    const newOnes = addedSet.filter((d) => !removedSet.includes(d));
    const goneOnes = removedSet.filter((d) => !addedSet.includes(d));
    const changedSignatures = addedSet.filter((d) => removedSet.includes(d));
    const modified = unique([...changedSignatures, ...touched]).filter((d) => !newOnes.includes(d) && !goneOnes.includes(d));

    let sentence: string;
    if (change.status === 'added') {
        sentence = `Se creó ${name}` + (newOnes.length ? ` con ${list(newOnes)}` : '');
    } else if (change.status === 'deleted') {
        sentence = `Se eliminó ${name}`;
    } else {
        const verb = change.status === 'renamed' ? `Se renombró ${path.posix.basename(change.oldPath)} a ${name}` : `Se modificó ${name}`;
        const parts: string[] = [];
        if (newOnes.length) parts.push(`se agrega${newOnes.length > 1 ? 'n' : ''} ${list(newOnes)}`);
        if (modified.length) parts.push(`se ajusta la lógica de ${list(modified)}`);
        if (goneOnes.length) parts.push(`se elimina${goneOnes.length > 1 ? 'n' : ''} ${list(goneOnes)}`);
        sentence = parts.length ? `${verb}: ${parts.join('; ')}` : `${verb} (+${additions}/-${deletions} líneas)`;
    }
    const newEndpoints = unique(endpoints);
    if (newEndpoints.length) {
        sentence += `; consume ${list(newEndpoints, 3)}`;
    }

    return {
        path: change.path,
        name,
        layer,
        status: change.status,
        additions,
        deletions,
        added: newOnes,
        removed: goneOnes,
        modified,
        endpoints: newEndpoints,
        sentence
    };
}

/**
 * Builds the "Diseño Técnico" paragraph: one sentence per file, grouped by layer, in the
 * style of the existing GEIR entries ("A nivel de backend se …").
 */
export function technicalDesign(changes: FileChange[]): { text: string; files: FileSummary[] } {
    const files = changes.map(summarizeFile);
    if (!files.length) {
        return { text: '', files };
    }
    const order: Layer[] = ['backend', 'base de datos', 'frontend', 'configuración'];
    const paragraphs: string[] = [];
    for (const layer of order) {
        const group = files.filter((f) => f.layer === layer);
        if (!group.length) continue;
        const intro = layer === 'configuración' ? 'En configuración' : `A nivel de ${layer}`;
        paragraphs.push(`${intro}: ${group.map((f) => f.sentence).join('. ')}.`);
    }
    return { text: paragraphs.join(' '), files };
}

