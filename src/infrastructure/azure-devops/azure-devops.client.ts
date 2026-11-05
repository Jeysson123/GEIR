import { BrowserContext, Page } from '@playwright/test';
import { Environment } from '../config/environment';

export interface DiscussionImage {
    url: string;
    base64: string;
    extension: string;
    author: string;
    date: string;
    /** The comment mentions "evidencia" / "capture", so it is the likely test evidence. */
    evidenceHint: boolean;
}

export interface WorkItemInfo {
    id: string;
    title: string;
    workItemType: string;
    state: string;
    descriptionHtml: string;
    acceptanceCriteriaHtml: string;
    iterationPath: string;
    editUrl: string;
    taskboardUrl: string;
    taskboardText: string;
    discussionImages: DiscussionImage[];
    warnings: string[];
}

export interface WorkItemSummary {
    id: string;
    title: string;
    workItemType: string;
    state: string;
    createdDate: string; // YYYY-MM-DD, '' when missing
    startDate: string;
    targetDate: string;
    closedDate: string;
    parentId: string;
    assignedTo: string;
    editUrl: string;
}

export class AzureAuthError extends Error {}

interface ApiResult<T> {
    status: number;
    data: T | null;
    message: string;
}

// Azure DevOps Server / TFS releases accept different REST versions (an unsupported one
// answers 400). '' sends no api-version, so the server uses its own default.
const API_VERSIONS = ['', '6.0', '5.1', '5.0', '4.1', '3.0', '2.0', '1.0'];
const COMMENT_VERSIONS = ['6.0-preview.3', '5.1-preview.3', '5.0-preview.2'];

/**
 * Reads work items from Azure DevOps Server with the user's network credentials. The
 * browser answers the Windows (NTLM/Negotiate) challenge, so requests go through page
 * navigations instead of Node HTTP.
 */
export class AzureDevOpsClient {
    private page: Page | null = null;
    private readonly base = Environment.services.azureDevOpsUrl.replace(/\/$/, '');
    private readonly collection = Environment.geir.azure.collection;
    private readonly project = Environment.geir.azure.project;

    constructor(private context: BrowserContext) {}

    private async getPage(): Promise<Page> {
        if (!this.page || this.page.isClosed()) {
            this.page = await this.context.newPage();
        }
        return this.page;
    }

    private async fetch(url: string): Promise<{ status: number; body: Buffer; contentType: string }> {
        const page = await this.getPage();
        // Edge aborts navigations that end in an empty error page (e.g. 401 after a failed
        // Windows auth) with ERR_HTTP_RESPONSE_CODE_FAILURE, so keep the real status here.
        let lastStatus = 0;
        const onResponse = (r: import('@playwright/test').Response) => {
            if (r.request().isNavigationRequest() && r.frame() === page.mainFrame()) {
                lastStatus = r.status();
            }
        };
        page.on('response', onResponse);
        try {
            const response = await page.goto(url, { waitUntil: 'commit', timeout: Environment.services.timeout });
            if (!response) {
                return { status: 0, body: Buffer.alloc(0), contentType: '' };
            }
            const body = await response.body().catch(() => Buffer.alloc(0));
            return { status: response.status(), body, contentType: response.headers()['content-type'] ?? '' };
        } catch (error: any) {
            const message = String(error?.message);
            if (/ERR_INVALID_AUTH_CREDENTIALS/.test(message)) {
                return { status: 401, body: Buffer.alloc(0), contentType: '' };
            }
            if (/ERR_HTTP_RESPONSE_CODE_FAILURE/.test(message) && lastStatus) {
                return { status: lastStatus, body: Buffer.alloc(0), contentType: '' };
            }
            if (/ERR_NAME_NOT_RESOLVED|ERR_CONNECTION|ERR_ADDRESS_UNREACHABLE|ERR_TIMED_OUT|Timeout/.test(String(error?.message))) {
                throw new Error(`No se pudo conectar a Azure DevOps (${this.base}). ¿Estás conectado a la VPN?`);
            }
            throw error;
        } finally {
            page.off('response', onResponse);
        }
    }

    private async json<T>(url: string): Promise<ApiResult<T>> {
        const res = await this.fetch(url);
        if (res.status === 401) {
            throw new AzureAuthError('Azure DevOps rechazó las credenciales de red.');
        }
        const text = res.body.toString('utf8');
        let parsed: any = null;
        try {
            parsed = JSON.parse(text);
        } catch {
            parsed = null;
        }
        if (res.status !== 200) {
            const message =
                parsed?.message ?? text.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 200);
            return { status: res.status, data: null, message };
        }
        return { status: 200, data: parsed as T | null, message: '' };
    }

    private workingVersion: string | null = null;

    /**
     * GET of a REST path, trying api-versions until the server accepts one. The version that
     * worked is tried first next time. Stops early on 404/403, which no version will fix.
     */
    private async api<T>(pathAndQuery: string, versions = API_VERSIONS): Promise<ApiResult<T>> {
        const shared = versions === API_VERSIONS;
        const ordered =
            shared && this.workingVersion !== null
                ? [this.workingVersion, ...versions.filter((v) => v !== this.workingVersion)]
                : versions;
        let last: ApiResult<T> = { status: 0, data: null, message: '' };
        for (const version of ordered) {
            const sep = pathAndQuery.includes('?') ? '&' : '?';
            last = await this.json<T>(`${this.base}${pathAndQuery}${version ? `${sep}api-version=${version}` : ''}`);
            if (last.status === 200 && last.data) {
                if (shared) {
                    this.workingVersion = version;
                }
                return last;
            }
            if (last.status === 404 || last.status === 403) {
                return last;
            }
        }
        return last;
    }

    /**
     * Confirms the credentials and returns the user's display name. Only a 401 means the
     * credentials were rejected; any other answer means Windows authentication succeeded.
     */
    async verify(): Promise<string> {
        const { status, data, message } = await this.api<any>(`/${this.collection}/_apis/connectionData?connectOptions=none`);
        if (status === 200 && data) {
            return data.authenticatedUser?.providerDisplayName ?? data.authenticatedUser?.customDisplayName ?? '';
        }
        if (status === 0) {
            throw new Error(`Azure DevOps no respondió (${this.base}).`);
        }
        console.warn(`[AZURE] connectionData respondió ${status}: ${message}. Credenciales aceptadas.`);
        return '';
    }

    editUrl(id: string): string {
        return `${this.base}/${this.collection}/${this.project}/_workitems/edit/${id}`;
    }

    async getWorkItem(id: string): Promise<WorkItemInfo> {
        const warnings: string[] = [];
        let fields: Record<string, any> | null = null;
        let relations: any[] = [];

        const item = await this.api<any>(`/${this.collection}/_apis/wit/workitems/${id}?$expand=relations`);
        if (item.status === 404) {
            throw new Error(`La tarea ${id} no existe en Azure DevOps.`);
        }
        if (item.data?.fields) {
            fields = item.data.fields;
            relations = item.data.relations ?? [];
        }

        let title = fields?.['System.Title'] ?? '';
        if (!fields) {
            console.warn(`[AZURE] workitems/${id} respondió ${item.status}: ${item.message}`);
            warnings.push(
                `La API de Azure DevOps respondió ${item.status} (${item.message}); se tomó el título desde la página de la tarea.`
            );
            title = await this.titleFromPage(id);
        }

        let descriptionHtml: string = fields?.['System.Description'] ?? '';
        const acceptanceCriteriaHtml: string = fields?.['Microsoft.VSTS.Common.AcceptanceCriteria'] ?? '';

        // Tasks often carry no description; their parent story does.
        if (!stripHtml(descriptionHtml)) {
            const parent = relations.find((r) => r.rel === 'System.LinkTypes.Hierarchy-Reverse');
            const parentId = parent?.url?.split('/').pop();
            if (parentId) {
                const parentFields = await this.fields(parentId);
                if (parentFields?.['System.Description']) {
                    descriptionHtml = parentFields['System.Description'];
                    warnings.push(`La descripción se tomó de la tarea padre ${parentId}.`);
                }
            }
        }

        const iterationPath: string = fields?.['System.IterationPath'] ?? '';
        const taskboard = this.taskboard(id, iterationPath);
        const discussionImages = await this.discussionImages(id).catch((error) => {
            warnings.push(`No se pudieron leer las imágenes de la discusión: ${error.message}`);
            return [] as DiscussionImage[];
        });

        return {
            id,
            title,
            workItemType: fields?.['System.WorkItemType'] ?? '',
            state: fields?.['System.State'] ?? '',
            descriptionHtml,
            acceptanceCriteriaHtml,
            iterationPath,
            editUrl: this.editUrl(id),
            taskboardUrl: taskboard.url,
            taskboardText: taskboard.text,
            discussionImages,
            warnings
        };
    }

    /**
     * What the admin options need from a work item: title, type, dates (for the work plan)
     * and parent (to relate a task to its user story). Null when it does not exist.
     */
    async getSummary(id: string): Promise<WorkItemSummary | null> {
        const item = await this.api<any>(`/${this.collection}/_apis/wit/workitems/${id}?$expand=relations`);
        const fields = item.data?.fields;
        if (!fields) {
            if (item.status === 404) {
                return null;
            }
            // Same fallback as getWorkItem: when the API does not answer, the story name
            // comes from its page so the sprint list and the e-mails never lose it.
            console.warn(`[AZURE] workitems/${id} respondió ${item.status}: ${item.message}`);
            const title = await this.titleFromPage(id).catch(() => '');
            if (!title) {
                return null;
            }
            return {
                id,
                title,
                workItemType: '',
                state: '',
                createdDate: '',
                startDate: '',
                targetDate: '',
                closedDate: '',
                parentId: '',
                assignedTo: '',
                editUrl: this.editUrl(id)
            };
        }
        const parent = (item.data.relations ?? []).find((r: any) => r.rel === 'System.LinkTypes.Hierarchy-Reverse');
        const date = (name: string) => (fields[name] ? String(fields[name]).slice(0, 10) : '');
        return {
            id,
            title: fields['System.Title'] ?? '',
            workItemType: fields['System.WorkItemType'] ?? '',
            state: fields['System.State'] ?? '',
            createdDate: date('System.CreatedDate'),
            startDate: date('Microsoft.VSTS.Scheduling.StartDate'),
            targetDate: date('Microsoft.VSTS.Scheduling.TargetDate') || date('Microsoft.VSTS.Scheduling.FinishDate'),
            closedDate: date('Microsoft.VSTS.Common.ClosedDate'),
            parentId: parent?.url?.split('/').pop() ?? '',
            assignedTo: fields['System.AssignedTo']?.displayName ?? String(fields['System.AssignedTo'] ?? '').replace(/\s*<.*>$/, ''),
            editUrl: this.editUrl(id)
        };
    }

    private async fields(id: string): Promise<Record<string, any> | null> {
        const { data } = await this.api<any>(`/${this.collection}/_apis/wit/workitems/${id}`);
        return data?.fields ?? null;
    }

    /** Same capture as tests/azure.spec.ts: the value of the "Title Field" input. */
    private async titleFromPage(id: string): Promise<string> {
        const page = await this.getPage();
        await page.goto(this.editUrl(id), { waitUntil: 'domcontentloaded' });
        const title = page.locator('input[aria-label="Title Field"]');
        await title.waitFor({ state: 'visible', timeout: Environment.services.timeout });
        return (await title.inputValue()).trim();
    }

    private taskboard(id: string, iterationPath: string): { url: string; text: string } {
        if (!iterationPath) {
            return { url: this.editUrl(id), text: this.editUrl(id) };
        }
        const team = Environment.geir.azure.team;
        const segments = iterationPath.split('\\').map(encodeURIComponent).join('/');
        const leaf = iterationPath.split('\\').pop() ?? iterationPath;
        return {
            url: `${this.base}/${this.collection}/${this.project}/_sprints/taskboard/${encodeURIComponent(team)}/${segments}?workitem=${id}`,
            text: `${team} ${leaf}`
        };
    }

    /** Comments of the Discussion section, newest first, as { html, author, date }. */
    private async comments(id: string): Promise<{ html: string; author: string; date: string }[]> {
        const commented = await this.api<any>(
            `/${this.collection}/${this.project}/_apis/wit/workItems/${id}/comments?$top=200`,
            COMMENT_VERSIONS
        );
        if (commented.data?.comments) {
            return commented.data.comments
                .map((c: any) => ({
                    html: c.text ?? c.renderedText ?? '',
                    author: c.createdBy?.displayName ?? c.revisedBy?.displayName ?? '',
                    date: c.createdDate ?? c.revisedDate ?? ''
                }))
                .sort((a: any, b: any) => (a.date < b.date ? 1 : -1));
        }
        // Older servers: the discussion lives in System.History of each revision.
        const updates = await this.api<any>(`/${this.collection}/_apis/wit/workitems/${id}/updates`);
        if (updates.data?.value) {
            return updates.data.value
                .filter((u: any) => u.fields?.['System.History']?.newValue)
                .map((u: any) => ({
                    html: u.fields['System.History'].newValue,
                    author: u.revisedBy?.displayName ?? '',
                    date: u.revisedDate ?? u.fields?.['System.ChangedDate']?.newValue ?? ''
                }))
                .reverse();
        }
        return [];
    }

    private async discussionImages(id: string): Promise<DiscussionImage[]> {
        const images: DiscussionImage[] = [];
        for (const comment of await this.comments(id)) {
            const hint = /evidencia|capture|captura/i.test(stripHtml(comment.html));
            for (const src of imageSources(comment.html)) {
                const url = new URL(src, `${this.base}/`).toString();
                const res = await this.fetch(url);
                if (res.status !== 200 || res.body.length === 0) {
                    continue;
                }
                const extension = extensionFor(res.contentType, url);
                if (!extension) {
                    continue;
                }
                images.push({
                    url,
                    base64: res.body.toString('base64'),
                    extension,
                    author: comment.author,
                    date: comment.date,
                    evidenceHint: hint
                });
            }
        }
        return images;
    }
}

export function stripHtml(html: string): string {
    return decodeEntities(
        html
            .replace(/<\s*br\s*\/?>/gi, '\n')
            .replace(/<\/(p|div|li|h\d|tr)>/gi, '\n')
            .replace(/<li[^>]*>/gi, '• ')
            .replace(/<[^>]+>/g, '')
    )
        .replace(/[ \t\u00a0]+/g, ' ')
        .replace(/\n\s*\n+/g, '\n')
        .trim();
}

function decodeEntities(text: string): string {
    const named: Record<string, string> = {
        amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ',
        aacute: 'á', eacute: 'é', iacute: 'í', oacute: 'ó', uacute: 'ú', ntilde: 'ñ',
        Aacute: 'Á', Eacute: 'É', Iacute: 'Í', Oacute: 'Ó', Uacute: 'Ú', Ntilde: 'Ñ', uuml: 'ü'
    };
    return text.replace(/&(#x?[0-9a-f]+|\w+);/gi, (all, code: string) => {
        if (code[0] === '#') {
            const n = code[1].toLowerCase() === 'x' ? parseInt(code.slice(2), 16) : parseInt(code.slice(1), 10);
            return Number.isFinite(n) ? String.fromCodePoint(n) : all;
        }
        return named[code] ?? all;
    });
}

function imageSources(html: string): string[] {
    const out: string[] = [];
    const re = /<img[^>]+src\s*=\s*["']([^"']+)["']/gi;
    let m: RegExpExecArray | null;
    while ((m = re.exec(html))) {
        if (!m[1].startsWith('data:')) {
            out.push(decodeEntities(m[1]));
        }
    }
    return out;
}

function extensionFor(contentType: string, url: string): string | null {
    if (/png/i.test(contentType)) return 'png';
    if (/jpe?g/i.test(contentType)) return 'jpeg';
    if (/gif/i.test(contentType)) return 'gif';
    const name = /fileName=([^&]+)/i.exec(url)?.[1] ?? url;
    const ext = /\.(png|jpe?g|gif)\b/i.exec(decodeURIComponent(name))?.[1]?.toLowerCase();
    return ext ? (ext === 'jpg' ? 'jpeg' : ext) : null;
}
