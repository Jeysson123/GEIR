import * as crypto from 'crypto';
import { APIRequestContext } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { Environment } from '../config/environment';

/** SharePoint did not accept the user's token (or there is none yet): sign in again. */
export class SessionExpiredError extends Error {
    constructor() {
        super('La sesión de Microsoft 365 expiró. Cierra sesión y vuelve a entrar.');
    }
}

export type DocumentLocation =
    | { kind: 'local'; filePath: string }
    | { kind: 'sharepoint'; siteUrl: string; serverRelativePath?: string; fileId?: string; shareUrl?: string };

export interface StoredDocument {
    data: Buffer | null; // null when the file does not exist yet
    name: string;
    display: string;
    webUrl: string;
}

const FILE_NAME = Environment.geir.fileName;

/**
 * Understands what the user pastes as "ruta de OneDrive": a folder view link
 * (…/shared?id=/personal/…), an Excel Online link (…Doc.aspx?sourcedoc={GUID}), a direct
 * file URL, a server-relative path, a sharing link (/:x:/g/…), or a local path.
 * A folder means "N - GEIR.xlsx inside it".
 */
export function resolveLocation(input: string): DocumentLocation {
    const raw = input.trim().replace(/^"|"$/g, '');
    if (!raw) {
        throw new Error('Escribe la ruta del documento.');
    }

    if (/^[a-zA-Z]:[\\/]/.test(raw) || raw.startsWith('\\\\') || raw.startsWith('./') || raw.startsWith('../')) {
        let filePath = path.resolve(raw);
        const isDir = fs.existsSync(filePath) && fs.statSync(filePath).isDirectory();
        if (isDir || !/\.xlsx$/i.test(filePath)) {
            filePath = path.join(filePath, FILE_NAME);
        }
        return { kind: 'local', filePath };
    }

    let url: URL;
    try {
        url = new URL(raw.startsWith('/') ? `https://${Environment.geir.sharepointHost}${raw}` : raw);
    } catch {
        throw new Error('La ruta no es un enlace de OneDrive/SharePoint ni una ruta local válida.');
    }
    if (!/sharepoint\.com$/i.test(url.hostname)) {
        throw new Error(`El enlace debe ser de SharePoint/OneDrive (${Environment.geir.sharepointHost}).`);
    }

    // Excel Online link: the file is identified by its GUID.
    const sourcedoc = url.searchParams.get('sourcedoc');
    if (sourcedoc) {
        const sitePath = url.pathname.replace(/^\/:[a-z]:\/r/i, '').split('/_layouts/')[0];
        return { kind: 'sharepoint', siteUrl: `${url.origin}${sitePath}`, fileId: sourcedoc.replace(/[{}]/g, '') };
    }

    // Folder view: …/shared?id=/personal/x/Documents/A/B or …/AllItems.aspx?id=…
    const id = url.searchParams.get('id');
    let serverRelativePath = id ?? decodeURIComponent(url.pathname.replace(/^\/:[a-z]:\/r/i, ''));

    // Sharing links (/:x:/g/…, /:f:/s/…) must be opened to learn the real path.
    if (!id && /^\/:[a-z]:\/[gs]\//i.test(url.pathname)) {
        return { kind: 'sharepoint', siteUrl: url.origin, shareUrl: url.toString() };
    }
    if (!id && serverRelativePath.includes('/_layouts/')) {
        throw new Error('No se reconoce el enlace. Pega el enlace de la carpeta o del archivo de Excel.');
    }

    serverRelativePath = serverRelativePath.replace(/\/+$/, '');
    if (!/\.xlsx$/i.test(serverRelativePath)) {
        serverRelativePath = `${serverRelativePath}/${FILE_NAME}`;
    }
    return { kind: 'sharepoint', siteUrl: `${url.origin}${siteOf(serverRelativePath)}`, serverRelativePath };
}

function siteOf(serverRelativePath: string): string {
    const m = /^\/(personal|sites|teams)\/[^/]+/i.exec(serverRelativePath);
    return m ? m[0] : '';
}

export function describeLocation(loc: DocumentLocation): string {
    if (loc.kind === 'local') {
        return loc.filePath;
    }
    return loc.serverRelativePath ? `${new URL(loc.siteUrl).origin}${loc.serverRelativePath}` : loc.siteUrl;
}

// ------------------------------------------------------------------------ local files

export function readLocal(loc: { filePath: string }): StoredDocument {
    const exists = fs.existsSync(loc.filePath);
    return {
        data: exists ? fs.readFileSync(loc.filePath) : null,
        name: path.basename(loc.filePath),
        display: loc.filePath,
        webUrl: ''
    };
}

export function writeLocal(loc: { filePath: string }, data: Buffer): StoredDocument {
    fs.mkdirSync(path.dirname(loc.filePath), { recursive: true });
    fs.writeFileSync(loc.filePath, data);
    return { data, name: path.basename(loc.filePath), display: loc.filePath, webUrl: '' };
}

// ------------------------------------------------------------------------ SharePoint

/** How SharePoint calls are authenticated: the user's own HTTP client and access token. */
export interface SharePointAuth {
    http: APIRequestContext;
    bearer: string;
}

/**
 * Reads and writes files through the SharePoint REST API from Node, with the access token
 * of the signed-in user. No browser page is involved.
 */
export class OneDriveFilesClient {
    /** @param auth how to authenticate, or null when a sign-in is needed. */
    constructor(private auth: () => Promise<SharePointAuth | null>) {}

    private async call(
        siteUrl: string,
        method: 'GET' | 'POST',
        apiPath: string,
        options: { body?: Buffer; binary?: boolean } = {}
    ): Promise<{ status: number; json?: any; data?: Buffer; text?: string }> {
        const auth = await this.auth();
        if (!auth) {
            throw new SessionExpiredError();
        }
        const headers: Record<string, string> = {
            Accept: 'application/json;odata=nometadata',
            Authorization: `Bearer ${auth.bearer}`
        };
        if (options.body) {
            headers['Content-Type'] = 'application/octet-stream';
        }
        const res = await auth.http.fetch(`${siteUrl}${apiPath}`, {
            method,
            headers,
            data: options.body ?? (method === 'POST' ? '' : undefined),
            maxRedirects: 0,
            // Uploads can be large (Deploy jars of ~40 MB over the VPN).
            ...(options.body ? { timeout: 15 * 60 * 1000 } : {})
        });
        const status = res.status();
        // Without a valid session SharePoint answers 401, 403 + X-Forms_Based_Auth_Required,
        // or redirects to the Microsoft login.
        if (status === 401 || (status === 403 && res.headers()['x-forms_based_auth_required']) || (status >= 300 && status < 400)) {
            throw new SessionExpiredError();
        }
        if (status === 403) {
            throw new Error(
                `SharePoint respondió 403: tu usuario no tiene permiso sobre ${siteUrl}. Verifica que la carpeta esté compartida contigo.`
            );
        }
        if (options.binary && status === 200) {
            return { status, data: await res.body() };
        }
        const text = await res.text();
        let json: any;
        try {
            json = JSON.parse(text);
        } catch {
            json = undefined;
        }
        return { status, json, text: json ? undefined : text.slice(0, 500) };
    }

    /** Turns GUID links and sharing links into a site + server-relative path. */
    async normalize(
        loc: Extract<DocumentLocation, { kind: 'sharepoint' }>
    ): Promise<{ siteUrl: string; serverRelativePath: string }> {
        if (loc.serverRelativePath) {
            return { siteUrl: loc.siteUrl, serverRelativePath: loc.serverRelativePath };
        }
        if (loc.shareUrl) {
            // Sharing link: ask SharePoint which file it points to.
            const encoded = 'u!' + Buffer.from(loc.shareUrl).toString('base64').replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
            const res = await this.call(loc.siteUrl, 'GET', `/_api/v2.1/shares/${encoded}/driveItem?$select=webUrl`);
            const resolved = res.json?.webUrl ? resolveLocation(res.json.webUrl) : null;
            if (!resolved || resolved.kind !== 'sharepoint' || resolved.shareUrl) {
                throw new Error(`No se pudo abrir el enlace compartido (${res.status}).`);
            }
            return this.normalize(resolved);
        }
        const res = await this.call(loc.siteUrl, 'GET', `/_api/web/GetFileById('${loc.fileId}')?$select=ServerRelativeUrl`);
        if (res.status !== 200 || !res.json?.ServerRelativeUrl) {
            throw new Error('No se encontró el archivo del enlace de Excel.');
        }
        return { siteUrl: loc.siteUrl, serverRelativePath: res.json.ServerRelativeUrl as string };
    }

    async read(loc: Extract<DocumentLocation, { kind: 'sharepoint' }>): Promise<StoredDocument & { serverRelativePath: string }> {
        const { siteUrl, serverRelativePath } = await this.normalize(loc);
        const fileApi = `/_api/web/GetFileByServerRelativePath(decodedurl='${odataPath(serverRelativePath)}')`;
        const meta = await this.call(siteUrl, 'GET', `${fileApi}?$select=Name,UniqueId,Exists`);
        const name = path.posix.basename(serverRelativePath);
        const display = `${new URL(siteUrl).origin}${serverRelativePath}`;
        if (meta.status === 404 || meta.json?.Exists === false) {
            return { data: null, name, display, webUrl: '', serverRelativePath };
        }
        if (meta.status !== 200) {
            throw new Error(`SharePoint respondió ${meta.status} al leer ${display}`);
        }
        const content = await this.call(siteUrl, 'GET', `${fileApi}/$value`, { binary: true });
        if (content.status !== 200 || !content.data) {
            throw new Error(`No se pudo descargar ${display} (${content.status})`);
        }
        return {
            data: content.data,
            name,
            display,
            webUrl: excelWebUrl(siteUrl, meta.json.UniqueId, name),
            serverRelativePath
        };
    }

    async write(loc: Extract<DocumentLocation, { kind: 'sharepoint' }>, data: Buffer): Promise<StoredDocument> {
        const { siteUrl, serverRelativePath } = await this.normalize(loc);
        const folder = path.posix.dirname(serverRelativePath);
        const name = path.posix.basename(serverRelativePath);
        await this.ensureFolder(siteUrl, folder);
        const res = await this.call(
            siteUrl,
            'POST',
            `/_api/web/GetFolderByServerRelativePath(decodedurl='${odataPath(folder)}')/Files/AddUsingPath(decodedurl='${odataPath(name)}',overwrite=true)`,
            // Large files (Deploy jars) go in parts: one request of 40 MB can be cut by the proxy.
            { body: data.length > UPLOAD_CHUNK ? Buffer.alloc(0) : data }
        );
        if (res.status === 423 || /locked|bloqueado/i.test(res.text ?? '')) {
            throw new Error('El archivo está bloqueado por otro usuario o aplicación. Ciérralo e intenta de nuevo.');
        }
        if (res.status !== 200) {
            throw new Error(`SharePoint respondió ${res.status} al guardar: ${res.text ?? JSON.stringify(res.json ?? {})}`);
        }
        if (data.length > UPLOAD_CHUNK) {
            await this.uploadInParts(siteUrl, serverRelativePath, data);
        }
        const meta = await this.call(
            siteUrl,
            'GET',
            `/_api/web/GetFileByServerRelativePath(decodedurl='${odataPath(serverRelativePath)}')?$select=UniqueId`
        );
        return {
            data,
            name,
            display: `${new URL(siteUrl).origin}${serverRelativePath}`,
            webUrl: excelWebUrl(siteUrl, meta.json?.UniqueId, name)
        };
    }

    /**
     * SharePoint's chunked upload (StartUpload → ContinueUpload → FinishUpload) on the file
     * just created empty.
     */
    private async uploadInParts(siteUrl: string, serverRelativePath: string, data: Buffer) {
        const file = `/_api/web/GetFileByServerRelativePath(decodedurl='${odataPath(serverRelativePath)}')`;
        const id = crypto.randomUUID();
        for (let offset = 0; offset < data.length; offset += UPLOAD_CHUNK) {
            const part = data.subarray(offset, Math.min(offset + UPLOAD_CHUNK, data.length));
            const first = offset === 0;
            const last = offset + part.length >= data.length;
            const action = first && last
                ? null
                : first
                ? `StartUpload(uploadId=guid'${id}')`
                : last
                ? `FinishUpload(uploadId=guid'${id}',fileOffset=${offset})`
                : `ContinueUpload(uploadId=guid'${id}',fileOffset=${offset})`;
            if (!action) {
                break;
            }
            const res = await this.call(siteUrl, 'POST', `${file}/${action}`, { body: Buffer.from(part) });
            if (res.status !== 200) {
                if (!first) {
                    await this.call(siteUrl, 'POST', `${file}/CancelUpload(uploadId=guid'${id}')`).catch(() => undefined);
                }
                throw new Error(
                    `SharePoint respondió ${res.status} al subir la parte ${Math.floor(offset / UPLOAD_CHUNK) + 1}: ${
                        res.text ?? JSON.stringify(res.json ?? {})
                    }`
                );
            }
        }
    }

    /** Size in bytes of a file in OneDrive, or null when it is not there. */
    async size(loc: Extract<DocumentLocation, { kind: 'sharepoint' }>): Promise<number | null> {
        const { siteUrl, serverRelativePath } = await this.normalize(loc);
        const res = await this.call(
            siteUrl,
            'GET',
            `/_api/web/GetFileByServerRelativePath(decodedurl='${odataPath(serverRelativePath)}')?$select=Length,Exists`
        );
        if (res.status !== 200 || res.json?.Exists === false) {
            return null;
        }
        return Number(res.json?.Length ?? NaN);
    }

    /** Deletes the file from OneDrive into the recycle bin (recoverable). */
    async recycle(loc: Extract<DocumentLocation, { kind: 'sharepoint' }>): Promise<void> {
        const { siteUrl, serverRelativePath } = await this.normalize(loc);
        const res = await this.call(
            siteUrl,
            'POST',
            `/_api/web/GetFileByServerRelativePath(decodedurl='${odataPath(serverRelativePath)}')/recycle()`
        );
        if (res.status === 423 || /locked|bloqueado/i.test(res.text ?? '')) {
            throw new Error('El archivo está abierto o bloqueado en OneDrive. Ciérralo en Excel e intenta de nuevo.');
        }
        if (res.status !== 200 && res.status !== 404) {
            throw new Error(`SharePoint respondió ${res.status} al eliminar el archivo: ${res.text ?? JSON.stringify(res.json ?? {})}`);
        }
    }

    private async ensureFolder(siteUrl: string, folder: string) {
        const sitePath = new URL(siteUrl).pathname.replace(/\/$/, '');
        const parts = folder.slice(sitePath.length).split('/').filter(Boolean);
        let current = sitePath;
        for (const part of parts) {
            current = `${current}/${part}`;
            const res = await this.call(siteUrl, 'GET', `/_api/web/GetFolderByServerRelativePath(decodedurl='${odataPath(current)}')?$select=Exists`);
            if (res.status === 200 && res.json?.Exists) {
                continue;
            }
            const created = await this.call(siteUrl, 'POST', `/_api/web/Folders/AddUsingPath(decodedurl='${odataPath(current)}')`);
            if (created.status !== 200) {
                throw new Error(`No se pudo crear la carpeta ${current} (${created.status})`);
            }
        }
    }
}

/** Parts of 8 MB for large uploads. */
export const UPLOAD_CHUNK = 8 * 1024 * 1024;

function odataPath(p: string): string {
    return p
        .replace(/'/g, "''")
        .split('/')
        .map(encodeURIComponent)
        .join('/');
}

function excelWebUrl(siteUrl: string, uniqueId: string | undefined, name: string): string {
    if (!uniqueId) {
        return '';
    }
    return `${siteUrl}/_layouts/15/Doc.aspx?sourcedoc=%7B${uniqueId.replace(/[{}]/g, '')}%7D&file=${encodeURIComponent(name)}&action=default&mobileredirect=true`;
}
