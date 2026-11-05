import { APIRequestContext, BrowserContext, request } from '@playwright/test';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { Environment } from '../infrastructure/config/environment';
import { AzureAuthError, AzureDevOpsClient, WorkItemInfo } from '../infrastructure/azure-devops/azure-devops.client';
import { GitLabClient, TaskChanges } from '../infrastructure/gitlab/gitlab.client';
import { MfaPrompt, MicrosoftCredentialsError, signInOnPage } from '../infrastructure/microsoft/microsoft-login';
import { DevicePrompt, MicrosoftToken } from '../infrastructure/microsoft/microsoft-token';
import { GRAPH, GraphAuthError, OutlookDraftsClient, SavedDraft } from '../infrastructure/microsoft/outlook-drafts.client';
import {
    DocumentLocation,
    OneDriveFilesClient,
    SessionExpiredError,
    SharePointAuth,
    StoredDocument,
    describeLocation,
    readLocal,
    resolveLocation,
    writeLocal
} from '../infrastructure/microsoft/onedrive-files.client';
import { GeirDocument, GeirTask, TaskSummary } from '../domain/document';
import { buildDraft } from '../domain/draft';
import { FileSummary } from '../domain/change-summary';
import { mockChanges, mockWorkItem } from './mock-data';
import { NetworkCredentials, Session, browser, createSession, destroySession } from './session';

export const MOCK = Environment.geir.mock;
const TEMPLATE = path.resolve(__dirname, '../..', Environment.geir.templatePath);

export class UserError extends Error {
    constructor(message: string, readonly status = 400) {
        super(message);
    }
}

// ------------------------------------------------------------------------ clients

const azureClients = new WeakMap<Session, AzureDevOpsClient>();
const gitlabClients = new WeakMap<Session, GitLabClient>();
const oneDriveClients = new WeakMap<Session, OneDriveFilesClient>();
const microsoftSignIns = new WeakMap<Session, Promise<void>>();
const microsoftTokens = new WeakMap<Session, MicrosoftToken>();

export async function azure(session: Session): Promise<AzureDevOpsClient> {
    let client = azureClients.get(session);
    if (!client) {
        const context = await session.context('azure', async () =>
            (await browser()).newContext({
                httpCredentials: { username: session.names.windows, password: session.credentials.password },
                ignoreHTTPSErrors: true
            })
        );
        client = new AzureDevOpsClient(context);
        azureClients.set(session, client);
        context.on('close', () => azureClients.delete(session));
    }
    return client;
}

async function gitlab(session: Session): Promise<GitLabClient> {
    let client = gitlabClients.get(session);
    if (!client) {
        // A browser context of this user only, kept in memory while the session lives.
        const context = await session.context('gitlab', async () => (await browser()).newContext({ ignoreHTTPSErrors: true }));
        client = new GitLabClient(context);
        if (!(await client.isSignedIn())) {
            await client.signIn(session.names.ldap, session.credentials.password);
        }
        gitlabClients.set(session, client);
        context.on('close', () => gitlabClients.delete(session));
    }
    return client;
}

/**
 * Microsoft 365 access for a session: a token for the SharePoint host, signed in with a
 * device code (no automated browser, which managed Edge on the VM closes after login).
 */
function microsoftToken(session: Session): MicrosoftToken {
    let token = microsoftTokens.get(session);
    if (!token) {
        // Only in memory and in the user's browser (sessionStorage): nothing on the server disk.
        const created = new MicrosoftToken(`https://${Environment.geir.sharepointHost}`, session.microsoftRefreshToken, (refresh) =>
            session.setMicrosoftRefreshToken(refresh)
        );
        session.onClose(() => created.dispose());
        microsoftTokens.set(session, created);
        token = created;
    }
    return token;
}

function oneDrive(session: Session): OneDriveFilesClient {
    let client = oneDriveClients.get(session);
    if (!client) {
        client = new OneDriveFilesClient(() => microsoftAuth(session));
        oneDriveClients.set(session, client);
    }
    return client;
}

/**
 * SharePoint access of this user: their own token and HTTP client. Several people share
 * the server, so there is no saved session on disk that everybody would use.
 */
async function microsoftAuth(session: Session): Promise<SharePointAuth | null> {
    const ms = microsoftToken(session);
    const token = await ms.token();
    return token ? { http: await ms.http(), bearer: token } : null;
}

/**
 * Signs in to Microsoft 365: GEIR shows a code (session.mfaPrompt, kind 'device') that the
 * user types at microsoft.com/devicelogin in their own browser, where they are already
 * signed in. The token is then renewed on its own, so this is needed rarely.
 */
function signInMicrosoft(session: Session): Promise<void> {
    let pending = microsoftSignIns.get(session);
    if (!pending) {
        const email = session.names.email;
        pending = (async () => {
            console.log(`[GEIR] Iniciando sesión en Microsoft 365 como ${email}`);
            let finished = false;
            let automation: Promise<void> | null = null;
            try {
                await microsoftToken(session).signIn((prompt) => {
                    session.mfaPrompt = {
                        kind: 'device',
                        number: prompt.userCode,
                        url: prompt.verificationUri,
                        message: `Abre ${prompt.verificationUri} en tu navegador, escribe este código y entra con ${email}.`,
                        auto: 'working',
                        autoMessage: 'GEIR está iniciando sesión en Microsoft por ti…'
                    };
                    automation = automateDeviceLogin(session, prompt, () => finished);
                });
                console.log('[GEIR] Sesión de Microsoft 365 lista');
            } catch (error: any) {
                throw new UserError(`Microsoft 365: ${error.message}`, 502);
            } finally {
                finished = true;
                session.mfaPrompt = null;
                await (automation as Promise<void> | null)?.catch(() => undefined);
            }
        })().finally(() => microsoftSignIns.delete(session));
        microsoftSignIns.set(session, pending);
    }
    return pending;
}

/**
 * Types the device code, e-mail and password in a private headless browser, so the user
 * only approves MFA in Authenticator (the number is shown in GEIR). The token itself is
 * received by Node (MicrosoftToken.signIn), so it does not matter if the browser is closed
 * afterwards. If anything fails, the code stays on screen to finish by hand.
 */
async function automateDeviceLogin(session: Session, prompt: DevicePrompt, isDone: () => boolean): Promise<void> {
    const update = (patch: Partial<MfaPrompt>) => {
        if (!isDone() && session.mfaPrompt?.kind === 'device') {
            session.mfaPrompt = { ...session.mfaPrompt, ...patch };
        }
    };
    let context: BrowserContext | null = null;
    try {
        context = await (await browser()).newContext({
            viewport: { width: 1000, height: 720 },
            // Windows seamless sign-on asks for the network account; offered to that host only.
            httpCredentials: {
                username: session.names.windows,
                password: session.credentials.password,
                origin: 'https://autologon.microsoftazuread-sso.com'
            }
        });
        await signInOnPage(await context.newPage(), {
            username: session.names.email,
            password: session.credentials.password,
            startUrl: prompt.verificationUri,
            deviceCode: prompt.userCode,
            isDone,
            staySignedIn: false,
            timeoutMs: 600000,
            onMfa: (mfa) => {
                if (mfa?.kind === 'approve') {
                    console.log(`[GEIR] MFA: aprueba en Authenticator ${mfa.number}`);
                    update({ auto: 'approve', approveNumber: mfa.number, autoMessage: mfa.message });
                } else if (mfa?.kind === 'code') {
                    update({
                        auto: 'manual',
                        approveNumber: '',
                        autoMessage: 'Microsoft pide un código de verificación: termina con el botón de abajo.'
                    });
                } else {
                    update({ auto: 'working', approveNumber: '', autoMessage: 'GEIR está iniciando sesión en Microsoft por ti…' });
                }
            }
        });
        if (isDone()) {
            console.log('[GEIR] Inicio de sesión de Microsoft hecho automáticamente');
        }
    } catch (error: any) {
        if (!isDone()) {
            const reason =
                error instanceof MicrosoftCredentialsError
                    ? `Microsoft rechazó el correo ${session.names.email} o la contraseña`
                    : String(error?.message ?? error).slice(0, 160);
            console.log(`[GEIR] No se pudo iniciar sesión automáticamente: ${reason}`);
            update({ auto: 'manual', approveNumber: '', autoMessage: `No se pudo hacer automáticamente (${reason}). Termina con el botón de abajo.` });
        }
    } finally {
        await context?.close().catch(() => undefined);
    }
}

export function microsoftStatus(session: Session) {
    return { signingIn: microsoftSignIns.has(session), mfa: session.mfaPrompt };
}

export async function withMicrosoft<T>(session: Session, action: (client: OneDriveFilesClient) => Promise<T>): Promise<T> {
    for (let attempt = 1; ; attempt++) {
        try {
            return await action(oneDrive(session));
        } catch (error: any) {
            if (!(error instanceof SessionExpiredError) || attempt >= 3) {
                throw error;
            }
            microsoftToken(session).forget();
            await signInMicrosoft(session);
        }
    }
}

/**
 * Saves an e-mail as a DRAFT in the Outlook mailbox of the person using GEIR (Microsoft
 * Graph, with the same Microsoft 365 sign-in used for OneDrive). It is never sent.
 */
export async function saveOutlookDraft(session: Session, mime: Buffer): Promise<SavedDraft> {
    if (MOCK) {
        return { id: `demo-${Date.now()}`, subject: '', webLink: '' };
    }
    for (let attempt = 1; ; attempt++) {
        const token = await microsoftToken(session).token(GRAPH);
        if (!token) {
            if (attempt >= 3) {
                throw new Error('no hay sesión de Microsoft 365');
            }
            await signInMicrosoft(session);
            continue;
        }
        try {
            return await new OutlookDraftsClient(token, await microsoftToken(session).http()).saveMime(mime);
        } catch (error) {
            if (!(error instanceof GraphAuthError) || attempt >= 2) {
                throw error;
            }
            console.log(`[GEIR] Outlook rechazó el token (${error.message}); se inicia sesión de nuevo.`);
            microsoftToken(session).forget();
        }
    }
}

// ------------------------------------------------------------------------ auth

export async function login(credentials: NetworkCredentials): Promise<Session> {
    if (!credentials.username?.trim() || !credentials.password) {
        throw new UserError('Escribe tu usuario y contraseña de red.');
    }
    const email = credentials.email?.trim() ?? '';
    if (!email && !MOCK) {
        // Each person signs in to Microsoft 365 with their own mailbox (never a shared default).
        throw new UserError(`Escribe tu correo de Microsoft (ej. nombre_apellido@${Environment.geir.microsoft.emailDomain || 'empresa.com'}).`);
    }
    if (email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
        throw new UserError(`El correo de Microsoft no es válido (ej. nombre_apellido@${Environment.geir.microsoft.emailDomain || 'empresa.com'}).`);
    }
    const session = createSession({
        username: credentials.username.trim(),
        password: credentials.password,
        email,
        microsoftRefreshToken: credentials.microsoftRefreshToken
    });
    if (MOCK) {
        return session;
    }
    try {
        const name = await (await azure(session)).verify();
        session.displayName = name || session.credentials.username;
        return session;
    } catch (error: any) {
        await destroySession(session.token);
        if (error instanceof AzureAuthError) {
            throw new UserError('Usuario o contraseña de red incorrectos.', 401);
        }
        throw new UserError(error?.message ?? String(error), 502);
    }
}

// ------------------------------------------------------------------------ documents

export async function readDocument(session: Session, location: DocumentLocation): Promise<StoredDocument> {
    if (location.kind === 'local') {
        return readLocal(location);
    }
    // OneDrive: se descarga a un archivo local, se lee y se borra la copia local.
    const online = await withMicrosoft(session, (client) => client.read(location));
    if (!online.data) {
        return online;
    }
    const localPath = workFile(session, online.name);
    try {
        fs.writeFileSync(localPath, online.data);
        return { ...online, data: fs.readFileSync(localPath) };
    } finally {
        fs.rmSync(localPath, { force: true });
    }
}

export async function writeDocument(session: Session, location: DocumentLocation, data: Buffer): Promise<StoredDocument> {
    if (location.kind === 'local') {
        return writeLocal(location, data);
    }
    // Solo para crear un archivo nuevo en OneDrive; las ediciones van por editOnline().
    return withMicrosoft(session, (client) => client.write(location, data));
}

/** A temporary file in this session's own work folder (removed when the session ends). */
function workFile(session: Session, name: string): string {
    fs.mkdirSync(session.workDir, { recursive: true });
    return path.join(session.workDir, `${Date.now()}-${path.basename(name)}`);
}

/**
 * Edits a OneDrive document as a local file:
 * 1. downloads it, 2. makes the change locally, 3. deletes it from OneDrive (to the
 * recycle bin), 4. uploads the local file again, 5. deletes the local file.
 * Deleting online waits until the local change is done, so a failed edit never loses the
 * file; if the upload fails, the local copy is kept and its path reported.
 */
export async function editOnline(
    session: Session,
    location: Extract<DocumentLocation, { kind: 'sharepoint' }>,
    edit: (data: Buffer | null) => Promise<Buffer>
): Promise<StoredDocument> {
    // 1. Descargar
    const online = await withMicrosoft(session, (client) => client.read(location));
    const localPath = workFile(session, online.name);
    let keepLocal = false;
    try {
        if (online.data) {
            fs.writeFileSync(localPath, online.data);
            console.log(`[GEIR] 1/5 Descargado ${online.display}`);
        }
        // 2. Modificar localmente
        fs.writeFileSync(localPath, await edit(online.data ? fs.readFileSync(localPath) : null));
        console.log('[GEIR] 2/5 Modificado localmente');
        // 3. Eliminar de OneDrive
        if (online.data) {
            await withMicrosoft(session, (client) => client.recycle(location));
            console.log('[GEIR] 3/5 Eliminado de OneDrive (papelera de reciclaje)');
        }
        // 4. Subir de nuevo
        keepLocal = true;
        let stored: StoredDocument;
        try {
            stored = await withMicrosoft(session, (client) => client.write(location, fs.readFileSync(localPath)));
        } catch (error: any) {
            throw new UserError(
                `No se pudo subir el archivo a OneDrive: ${error.message}. Tu cambio quedó en ${localPath}` +
                    (online.data ? ' y el original está en la papelera de reciclaje de OneDrive.' : '.'),
                502
            );
        }
        console.log('[GEIR] 4/5 Subido a OneDrive');
        keepLocal = false;
        return stored;
    } finally {
        // 5. Borrar el archivo local
        if (!keepLocal) {
            fs.rmSync(localPath, { force: true });
            console.log('[GEIR] 5/5 Archivo local borrado');
        }
    }
}

export function locate(input: string): DocumentLocation {
    let location: DocumentLocation;
    try {
        location = resolveLocation(input);
    } catch (error: any) {
        throw new UserError(error.message);
    }
    if (location.kind === 'local' && !Environment.geir.allowLocalPaths && !MOCK) {
        throw new UserError('En el servidor solo se aceptan enlaces de OneDrive / SharePoint, no rutas locales.');
    }
    return location;
}

export async function loadOrCreate(session: Session, location: DocumentLocation) {
    let stored = await readDocument(session, location);
    let created = false;
    if (!stored.data) {
        // "Si el archivo no existe creas el archivo N - GEIR.xlsx"
        stored = await writeDocument(session, location, fs.readFileSync(TEMPLATE));
        created = true;
    }
    return { stored, doc: await loadDocument(stored.data!), created };
}

export async function loadDocument(data: Buffer): Promise<GeirDocument> {
    try {
        return await GeirDocument.load(data);
    } catch (error: any) {
        throw new UserError(`El archivo no tiene el formato GEIR esperado: ${error.message}`);
    }
}

export interface DocumentInfo {
    location: string;
    name: string;
    webUrl: string;
    created: boolean;
    tasks: TaskSummary[];
}

export async function openDocument(session: Session, input: string): Promise<DocumentInfo> {
    const location = locate(input);
    const { stored, doc, created } = await loadOrCreate(session, location);
    return {
        location: describeLocation(location),
        name: stored.name,
        webUrl: stored.webUrl,
        created,
        tasks: await doc.listTasks()
    };
}

// ------------------------------------------------------------------------ tasks

export interface LookupResult {
    mode: 'new' | 'edit';
    task: GeirTask;
    workItem: Omit<WorkItemInfo, 'discussionImages' | 'descriptionHtml' | 'acceptanceCriteriaHtml'> | null;
    images: { base64: string; extension: string; author: string; date: string; evidenceHint: boolean }[];
    files: FileSummary[];
    sources: TaskChanges['sources'];
    warnings: string[];
}

function checkId(id: string): string {
    const clean = String(id ?? '').trim().replace(/^#/, '');
    if (!/^\d{3,}$/.test(clean)) {
        throw new UserError('El ID de la tarea debe ser numérico (ej. 100007).');
    }
    return clean;
}

/**
 * New ID: drafts every sheet from Azure DevOps + GitLab. Existing ID: returns what the
 * document has so it can be edited; `refresh` re-drafts it from the sources instead.
 */
export async function lookupTask(session: Session, input: string, rawId: string, refresh = false): Promise<LookupResult> {
    const id = checkId(rawId);
    const { doc } = await loadOrCreate(session, locate(input));
    const existing = await doc.readTask(id);

    if (existing && !refresh) {
        return { mode: 'edit', task: existing, workItem: null, images: [], files: [], sources: [], warnings: [] };
    }

    const warnings: string[] = [];
    let workItem: WorkItemInfo;
    try {
        workItem = MOCK ? mockWorkItem(id) : await (await azure(session)).getWorkItem(id);
    } catch (error: any) {
        throw new UserError(error.message, 502);
    }
    warnings.push(...workItem.warnings);

    let changes: TaskChanges | null = null;
    try {
        changes = MOCK ? mockChanges(id) : await (await gitlab(session)).findChanges(id);
        warnings.push(...changes.warnings);
    } catch (error: any) {
        warnings.push(`GitLab: ${error.message}`);
    }

    const { task, files } = buildDraft(id, workItem, changes);
    if (existing) {
        task.number = existing.number;
        task.evidence = task.evidence ?? existing.evidence;
    }
    const { discussionImages, descriptionHtml, acceptanceCriteriaHtml, ...summary } = workItem;
    return {
        mode: existing ? 'edit' : 'new',
        task,
        workItem: summary,
        images: discussionImages.map(({ url, ...img }) => img),
        files,
        sources: changes?.sources ?? [],
        warnings
    };
}

export async function saveTask(
    session: Session,
    input: string,
    task: GeirTask,
    evidenceChanged: boolean
): Promise<{ result: 'added' | 'updated'; number: number | null; document: DocumentInfo }> {
    const id = checkId(task?.id);
    if (!task.title?.trim()) {
        throw new UserError('La tarea necesita un título.');
    }
    const location = locate(input);
    const toWrite: GeirTask = { ...task, id, evidence: evidenceChanged ? task.evidence : null };
    let doc!: GeirDocument;
    let created = false;
    let result!: 'added' | 'updated';
    let stored: StoredDocument;
    if (location.kind === 'local') {
        // Re-read right before writing so edits made by others meanwhile are kept.
        ({ doc, created } = await loadOrCreate(session, location));
        result = await doc.upsertTask(toWrite);
        stored = await writeDocument(session, location, await doc.save());
    } else {
        stored = await editOnline(session, location, async (data) => {
            created = !data;
            doc = await loadDocument(data ?? fs.readFileSync(TEMPLATE));
            result = await doc.upsertTask(toWrite);
            return doc.save();
        });
    }
    return {
        result,
        number: toWrite.number,
        document: {
            location: describeLocation(location),
            name: stored.name,
            webUrl: stored.webUrl,
            created,
            tasks: await doc.listTasks()
        }
    };
}

/** Removes the task (requerimiento) from the four sheets of the document. */
export async function deleteTask(session: Session, input: string, rawId: string): Promise<{ document: DocumentInfo }> {
    const id = checkId(rawId);
    const location = locate(input);
    let doc!: GeirDocument;
    let found = false;
    const remove = async (target: GeirDocument) => {
        doc = target;
        found = await target.deleteTask(id);
        if (!found) {
            throw new UserError(`La tarea ${id} no está en el documento.`, 404);
        }
        return target.save();
    };
    let stored: StoredDocument;
    if (location.kind === 'local') {
        const loaded = await loadOrCreate(session, location);
        stored = await writeDocument(session, location, await remove(loaded.doc));
    } else {
        stored = await editOnline(session, location, async (data) => {
            if (!data) {
                throw new UserError(`La tarea ${id} no está en el documento.`, 404);
            }
            return remove(await loadDocument(data));
        });
    }
    console.log(`[GEIR] Tarea ${id} eliminada del documento`);
    return {
        document: {
            location: describeLocation(location),
            name: stored.name,
            webUrl: stored.webUrl,
            created: false,
            tasks: await doc.listTasks()
        }
    };
}
