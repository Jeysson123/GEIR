import { APIResponse, BrowserContext } from '@playwright/test';
import { Environment } from '../config/environment';

export type ChangeStatus = 'added' | 'modified' | 'deleted' | 'renamed';

export interface FileChange {
    path: string;
    oldPath: string;
    status: ChangeStatus;
    diff: string;
    project: string;
}

export interface ChangeSource {
    kind: 'merge_request' | 'commit' | 'branch';
    project: string;
    ref: string;
    title: string;
    url: string;
}

export interface TaskChanges {
    sources: ChangeSource[];
    files: FileChange[];
    warnings: string[];
}

/**
 * Finds what was changed for an Azure task in GitLab: merge requests and commits whose
 * title/message mention the ID, and branches named after it. Uses the session cookie of
 * a signed-in browser context, so the REST API and raw .diff routes both work.
 */
export class GitLabClient {
    private readonly base = Environment.services.gitlabUrl.replace(/\/$/, '');

    constructor(private context: BrowserContext) {}

    private api(path: string): Promise<APIResponse> {
        return this.context.request.get(`${this.base}/api/v4${path}`, { timeout: Environment.services.timeout });
    }

    private async apiJson<T>(path: string): Promise<T | null> {
        try {
            const res = await this.api(path);
            return res.ok() ? ((await res.json()) as T) : null;
        } catch {
            return null;
        }
    }

    async isSignedIn(): Promise<boolean> {
        try {
            const res = await this.api('/user');
            return res.ok();
        } catch (error: any) {
            if (/ENOTFOUND|ECONNREFUSED|ETIMEDOUT|EAI_AGAIN/.test(String(error?.message))) {
                throw new Error(`No se pudo conectar a GitLab (${this.base}). ¿Estás conectado a la VPN?`);
            }
            return false;
        }
    }

    /** LDAP sign-in with the user's network account (GitLab login form). */
    async signIn(username: string, password: string) {
        const page = await this.context.newPage();
        try {
            await page.goto(`${this.base}/users/sign_in`, { waitUntil: 'domcontentloaded' });
            await page.waitForSelector('#ldapmain_username', { timeout: 30000 });
            await page.fill('#ldapmain_username', username);
            await page.fill('#ldapmain_password', password);
            await page.click('button[data-testid="sign-in-button"]');
            await page.waitForLoadState('domcontentloaded');
            await page.waitForTimeout(1000);
            if (page.url().includes('/users/sign_in')) {
                throw new Error('GitLab rechazó las credenciales de red.');
            }
        } finally {
            await page.close();
        }
    }

    async findChanges(taskId: string): Promise<TaskChanges> {
        const result: TaskChanges = { sources: [], files: [], warnings: [] };
        const files = new Map<string, FileChange>();
        const add = (change: FileChange) => {
            const key = `${change.project}:${change.path}`;
            const prev = files.get(key);
            if (!prev) {
                files.set(key, { ...change });
                return;
            }
            prev.diff = `${prev.diff}\n${change.diff}`;
            if (change.status === 'deleted' || prev.status === 'deleted') {
                prev.status = change.status;
            } else if (prev.status !== 'added') {
                prev.status = change.status === 'added' ? 'added' : prev.status;
            }
        };
        const mentions = (text: string) => new RegExp(`(^|\\D)${taskId}(\\D|$)`).test(text ?? '');

        for (const project of Environment.geir.gitlab.projects) {
            const enc = encodeURIComponent(project);
            const info = await this.apiJson<any>(`/projects/${enc}`);
            if (!info) {
                result.warnings.push(`No se pudo leer el proyecto ${project} por la API; se usa la búsqueda web.`);
                await this.webFallback(project, taskId, result, add);
                continue;
            }
            const seenCommits = new Set<string>();

            // 1. Merge requests that mention the task.
            const mrs =
                (await this.apiJson<any[]>(
                    `/projects/${enc}/merge_requests?search=${taskId}&in=title,description&state=all&per_page=20`
                )) ?? [];
            for (const mr of mrs.filter((m) => m.state !== 'closed' && (mentions(m.title) || mentions(m.description) || mentions(m.source_branch)))) {
                result.sources.push({
                    kind: 'merge_request',
                    project,
                    ref: `!${mr.iid}`,
                    title: mr.title,
                    url: mr.web_url
                });
                for (const d of await this.mergeRequestDiffs(enc, mr.iid)) {
                    add(toChange(d, project));
                }
                const mrCommits = (await this.apiJson<any[]>(`/projects/${enc}/merge_requests/${mr.iid}/commits?per_page=100`)) ?? [];
                mrCommits.forEach((c) => seenCommits.add(c.id));
            }

            // 2. Commits whose message mentions the task (outside those merge requests). The
            // search API needs GitLab advanced search; without it, scan recent commits of
            // every branch instead.
            let commits = (await this.apiJson<any[]>(`/projects/${enc}/search?scope=commits&search=${taskId}&per_page=50`)) ?? [];
            commits = commits.filter((c) => mentions(`${c.title}\n${c.message}`));
            if (!commits.length) {
                commits = await this.scanCommits(enc, mentions);
            }
            for (const commit of commits) {
                if (seenCommits.has(commit.id) || !mentions(`${commit.title}\n${commit.message}`)) {
                    continue;
                }
                if ((commit.parent_ids ?? []).length > 1 || /^Merge (branch|remote-tracking)/.test(commit.title ?? '')) {
                    continue;
                }
                seenCommits.add(commit.id);
                result.sources.push({
                    kind: 'commit',
                    project,
                    ref: commit.short_id ?? commit.id.slice(0, 8),
                    title: commit.title,
                    url: commit.web_url ?? `${this.base}/${project}/-/commit/${commit.id}`
                });
                for (const d of await this.paged(`/projects/${enc}/repository/commits/${commit.id}/diff`)) {
                    add(toChange(d, project));
                }
            }

            // 3. Branches named after the task, compared with the default branch.
            if (!result.sources.some((s) => s.project === project)) {
                const branches = (await this.apiJson<any[]>(`/projects/${enc}/repository/branches?search=${taskId}`)) ?? [];
                for (const branch of branches.filter((b) => mentions(b.name))) {
                    const cmp = await this.apiJson<any>(
                        `/projects/${enc}/repository/compare?from=${encodeURIComponent(info.default_branch)}&to=${encodeURIComponent(branch.name)}`
                    );
                    if (!cmp?.diffs?.length) {
                        continue;
                    }
                    result.sources.push({
                        kind: 'branch',
                        project,
                        ref: branch.name,
                        title: `Rama ${branch.name} vs ${info.default_branch}`,
                        url: branch.web_url ?? `${this.base}/${project}/-/tree/${encodeURIComponent(branch.name)}`
                    });
                    cmp.diffs.forEach((d: any) => add(toChange(d, project)));
                }
            }
        }

        result.files = [...files.values()];
        if (!result.sources.length) {
            result.warnings.push(
                `No se encontraron merge requests, commits ni ramas que mencionen ${taskId} en: ${Environment.geir.gitlab.projects.join(', ')}. ` +
                    'El ID debe estar en el mensaje del commit, en el título del merge request o en el nombre de la rama.'
            );
        }
        return result;
    }

    /**
     * Commits of the last year, on all branches, whose message mentions the task. Uses
     * ?all=true; on old GitLab versions that ignore it, walks the recently updated branches.
     */
    private async scanCommits(enc: string, mentions: (text: string) => boolean): Promise<any[]> {
        const since = new Date(Date.now() - 365 * 24 * 3600 * 1000).toISOString();
        const found = new Map<string, any>();
        const collect = async (query: string, pages: number) => {
            for (let page = 1; page <= pages; page++) {
                const batch = await this.apiJson<any[]>(
                    `/projects/${enc}/repository/commits?${query}&since=${encodeURIComponent(since)}&per_page=100&page=${page}`
                );
                if (!batch?.length) {
                    break;
                }
                batch.filter((c) => mentions(`${c.title}\n${c.message}`)).forEach((c) => found.set(c.id, c));
                if (batch.length < 100) {
                    break;
                }
            }
        };
        await collect('all=true', 30);
        if (!found.size) {
            const recent = Date.now() - 120 * 24 * 3600 * 1000;
            const branches = (await this.paged(`/projects/${enc}/repository/branches`))
                .filter((b) => !b.commit?.committed_date || Date.parse(b.commit.committed_date) > recent)
                .slice(0, 80);
            for (const branch of branches) {
                await collect(`ref_name=${encodeURIComponent(branch.name)}`, 3);
            }
        }
        return [...found.values()];
    }

    private async mergeRequestDiffs(enc: string, iid: number): Promise<any[]> {
        const diffs = await this.paged(`/projects/${enc}/merge_requests/${iid}/diffs`);
        if (diffs.length) {
            return diffs;
        }
        // GitLab < 15.7
        const changes = await this.apiJson<any>(`/projects/${enc}/merge_requests/${iid}/changes`);
        return changes?.changes ?? [];
    }

    private async paged(path: string): Promise<any[]> {
        const out: any[] = [];
        for (let page = 1; page <= 10; page++) {
            const sep = path.includes('?') ? '&' : '?';
            const batch = await this.apiJson<any[]>(`${path}${sep}per_page=100&page=${page}`);
            if (!batch?.length) {
                break;
            }
            out.push(...batch);
            if (batch.length < 100) {
                break;
            }
        }
        return out;
    }

    /** Web search of commits plus raw .diff pages, for when the API is not reachable. */
    private async webFallback(
        project: string,
        taskId: string,
        result: TaskChanges,
        add: (c: FileChange) => void
    ) {
        const page = await this.context.newPage();
        try {
            await page.goto(`${this.base}/${project}/-/commits/HEAD?search=${taskId}`, { waitUntil: 'domcontentloaded' });
            const shas = await page.$$eval('a[href*="/-/commit/"]', (links) =>
                Array.from(
                    new Set(
                        links
                            .map((a) => (a as HTMLAnchorElement).href.match(/\/-\/commit\/([0-9a-f]{7,40})/)?.[1])
                            .filter(Boolean) as string[]
                    )
                )
            );
            for (const sha of shas.slice(0, 30)) {
                const res = await this.context.request.get(`${this.base}/${project}/-/commit/${sha}.diff`);
                if (!res.ok()) {
                    continue;
                }
                result.sources.push({ kind: 'commit', project, ref: sha.slice(0, 8), title: sha, url: `${this.base}/${project}/-/commit/${sha}` });
                parseUnifiedDiff(await res.text()).forEach((d) => add({ ...d, project }));
            }
        } catch (error: any) {
            result.warnings.push(`Búsqueda web en ${project} falló: ${error.message}`);
        } finally {
            await page.close();
        }
    }
}

function toChange(d: any, project: string): FileChange {
    const status: ChangeStatus = d.new_file ? 'added' : d.deleted_file ? 'deleted' : d.renamed_file ? 'renamed' : 'modified';
    return { path: d.new_path ?? d.old_path, oldPath: d.old_path ?? d.new_path, status, diff: d.diff ?? '', project };
}

/** Splits a `git diff` text into per-file changes. */
export function parseUnifiedDiff(text: string): Omit<FileChange, 'project'>[] {
    const files: Omit<FileChange, 'project'>[] = [];
    for (const chunk of text.split(/^diff --git /m).slice(1)) {
        const header = /^a\/(.+?) b\/(.+?)$/m.exec(chunk);
        if (!header) {
            continue;
        }
        const status: ChangeStatus = /^new file mode/m.test(chunk)
            ? 'added'
            : /^deleted file mode/m.test(chunk)
            ? 'deleted'
            : /^rename from/m.test(chunk)
            ? 'renamed'
            : 'modified';
        const hunkStart = chunk.search(/^@@/m);
        files.push({
            oldPath: header[1],
            path: header[2],
            status,
            diff: hunkStart === -1 ? '' : chunk.slice(hunkStart)
        });
    }
    return files;
}
