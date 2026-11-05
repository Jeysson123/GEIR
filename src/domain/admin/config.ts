import type { BuildSettings } from '../build/projects';
import * as fs from 'fs';
import * as path from 'path';

/**
 * Admin options configuration (config/geir.config.json). Read on every use so edits apply
 * without restarting. GEIR_CONFIG_DIR points to another folder (e.g. on a deploy server).
 */
export const CONFIG_DIR = path.resolve(process.env.GEIR_CONFIG_DIR || path.resolve(__dirname, '../../../config'));

export interface DerivedFile {
    clave: string;
    hoja: string;
    archivo: string;
}

export interface PlanPhase {
    actividad: string;
    dias?: number;
    desarrollo?: boolean;
    minimo?: number;
}

export interface ApprovalEmail {
    clave: string;
    archivo: string;
    asunto: string;
    documento: string;
    adjunto: string;
    para: string[];
    cc: string[];
}

export interface ForumEmail {
    archivo: string;
    asunto: string;
    para: string[];
    cc: string[];
    saludo: string;
    estatusGeir: string;
    hora: string;
    ingenieroResponsable: string;
    especialista: string;
    etiquetaEspecialista?: string;
    usuarioDueno: string;
    usuarioCertifica: string;
    recursosBackend: string[];
    recursosFrontend: string[];
}

export interface AdminConfig {
    sistema: string;
    valorTemporal: string;
    tiposDeTrabajo: Record<string, string>;
    archivos: { geir: string; sprint: string; plantillaSprint?: string; formatoDescripcion?: string; derivados: DerivedFile[] };
    estimaciones: {
        fechaPrimeraHistoria: string[];
        fechaUltimaHistoria: string[];
        feriados: string[];
        fases: PlanPhase[];
    };
    personas: Record<string, string>;
    correos: {
        plantillas?: string; // folder with the Outlook HTML templates (config/correos)
        programacion: string;
        firma: { nombre: string; area?: string; cargo?: string; telefono?: string; lema: string; logo: string };
        aprobaciones: ApprovalEmail[];
        foro: ForumEmail;
    };
    /** Compilation of the binaries for Deploy (optional: defaults in domain/build/projects). */
    compilacion?: Partial<BuildSettings>;
}

/**
 * geir.config.json (in the repository, no personal data) with geir.config.local.json on top
 * when it exists: the real names, e-mail addresses and signature of each installation. The
 * local file is git-ignored, so people's data never reaches the public repository.
 */
export function loadAdminConfig(): AdminConfig {
    const read = (name: string, required: boolean) => {
        const file = path.join(CONFIG_DIR, name);
        if (!required && !fs.existsSync(file)) {
            return {};
        }
        try {
            return JSON.parse(fs.readFileSync(file, 'utf8'));
        } catch (error: any) {
            throw new Error(`No se pudo leer la configuración ${file}: ${error.message}`);
        }
    };
    return mergeConfig(read('geir.config.json', true), read('geir.config.local.json', false)) as AdminConfig;
}

/** Objects are merged key by key; arrays and plain values of `override` replace `base`. */
export function mergeConfig(base: any, override: any): any {
    if (!override || typeof override !== 'object' || Array.isArray(override)) {
        return override === undefined ? base : override;
    }
    if (!base || typeof base !== 'object' || Array.isArray(base)) {
        return override;
    }
    const out: Record<string, any> = { ...base };
    for (const [key, value] of Object.entries(override)) {
        out[key] = mergeConfig(base[key], value);
    }
    return out;
}

/** A file under the config folder (e.g. assets/logo.png), or null when missing. */
export function configAsset(relative: string): Buffer | null {
    const file = path.resolve(CONFIG_DIR, relative);
    if (!file.startsWith(CONFIG_DIR) || !fs.existsSync(file)) {
        return null;
    }
    return fs.readFileSync(file);
}

/** "Bug" → "Corrección Defecto"; anything else → the '*' value. */
export function workTypeLabel(config: AdminConfig, azureType: string): string {
    const types = config.tiposDeTrabajo ?? {};
    return types[azureType] ?? types['*'] ?? 'Funcionalidad';
}

/** Replaces {name} placeholders; unknown ones are left as they are. */
export function fill(template: string, values: Record<string, string>): string {
    return template.replace(/\{(\w+)\}/g, (all, key) => (key in values ? values[key] : all));
}

/** Names typed in the admin card: empty means the configured "TEMPORAL". */
export function orTemporal(config: AdminConfig, value: string | undefined): string {
    const clean = (value ?? '').trim();
    return clean || config.valorTemporal || 'TEMPORAL';
}

/** File names must not carry characters Windows / SharePoint reject. */
export function safeFileName(name: string): string {
    return name.replace(/[\\/:*?"<>|#%]+/g, '-').replace(/\s+/g, ' ').trim();
}
