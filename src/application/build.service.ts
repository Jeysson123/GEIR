import * as crypto from 'crypto';
import * as fs from 'fs';
import * as path from 'path';
import { Environment } from '../infrastructure/config/environment';
import { DocumentLocation, SessionExpiredError, describeLocation } from '../infrastructure/microsoft/onedrive-files.client';
import { Commit, GitRepo } from '../infrastructure/build/git-repo';
import { run } from '../infrastructure/build/shell';
import { loadAdminConfig } from '../domain/admin/config';
import {
    BuildProject,
    BuildSettings,
    DEFAULT_BUILD,
    Pom,
    deployRoot,
    expectedJar,
    RepoView,
    frontScript,
    mentionsAny,
    pickJar,
    planBuild,
    projectOf,
    readmeCommand,
    validBranch
} from '../domain/build/projects';
import { parseRequirementPairs } from '../domain/admin/requirements';
import { MOCK, UserError, azure, loadOrCreate, locate, withMicrosoft } from './geir.service';
import { Session } from './session';

/**
 * Admin option "Rama": updates the source repository to the branch, compiles the
 * projects its commits touched (Maven jars and the front end with `npm run buildprod`)
 * and copies the binaries into the Deploy folder next to the GEIR, with the same layout
 * the deploys use (Deploy/backend/..., Deploy/frontend/dist/js).
 *
 * Compiling takes minutes, so it runs as a job the page follows. The clone is shared by
 * every user of the server: builds run one at a time.
 */

type ProjectState = 'pendiente' | 'compilando' | 'listo' | 'error' | 'omitido';

export interface BuildJobView {
    id: string;
    branch: string;
    base: string;
    state: 'en cola' | 'trabajando' | 'listo' | 'error';
    step: string;
    startedAt: string;
    finishedAt: string;
    /** IDs of the GEIR (tasks, their stories, the requirement stories) searched in the commits. */
    ids: string[];
    /** Commits of the branch; `related` lists the GEIR IDs each one names (empty: not compiled). */
    commits: (Omit<Commit, 'message' | 'files'> & { related: string[]; projects: string[] })[];
    projects: {
        name: string;
        dir: string;
        kind: 'maven' | 'front';
        reason: BuildProject['reason'];
        destination: string;
        command: string;
        state: ProjectState;
        outputs: string[];
    }[];
    files: { path: string; size: number; error?: string }[];
    deploy: string;
    warnings: string[];
    error: string;
    log: string[];
}

interface BuildJob extends BuildJobView {
    owner: string;
}

const MAX_LOG = 3000;
const jobs = new Map<string, BuildJob>();
let queue: Promise<void> = Promise.resolve();

export function buildSettings(): BuildSettings {
    const custom = loadAdminConfig().compilacion ?? {};
    return {
        ramaBase: custom.ramaBase ?? DEFAULT_BUILD.ramaBase,
        recompilarDependientes: custom.recompilarDependientes ?? DEFAULT_BUILD.recompilarDependientes,
        maven: { ...DEFAULT_BUILD.maven, ...(custom.maven ?? {}) },
        front: { ...DEFAULT_BUILD.front, ...(custom.front ?? {}) }
    };
}

/** Starts (queues) the compilation of a branch for the GEIR at `location`. */
/**
 * The local source repository to use, first match wins:
 *   1. the folder typed in the panel (only where local paths are allowed);
 *   2. GEIR_REPO_DIR;
 *   3. a GEIR_REPO_NAME folder next to GEIR (C:\dev\GEIR → C:\dev\<name>), or the repo inside it.
 * Null: no local copy, GEIR clones GitLab.
 */
export function sourceFolder(typed = ''): string | null {
    const candidates: string[] = [];
    if (typed.trim()) {
        candidates.push(typed.trim());
    } else if (Environment.build.localRepo) {
        candidates.push(Environment.build.localRepo);
    } else {
        const roots = [process.cwd(), path.resolve(__dirname, '../..'), path.resolve(__dirname, '../../..')];
        for (const root of roots) {
            const name = Environment.build.repoName;
            if (name) {
                candidates.push(path.resolve(root, '..', name), path.resolve(root, '..', name.toLowerCase()));
            }
        }
    }
    for (const dir of [...new Set(candidates)]) {
        const found = gitRoot(dir);
        if (found) {
            return found;
        }
    }
    return null;
}

/** The folder itself when it is a clone, or its only (or GEIR_REPO_NAME) subfolder that is one. */
function gitRoot(dir: string): string | null {
    if (!fs.existsSync(dir)) {
        return null;
    }
    if (fs.existsSync(path.join(dir, '.git'))) {
        return dir;
    }
    let children: string[] = [];
    try {
        children = fs
            .readdirSync(dir, { withFileTypes: true })
            .filter((e) => e.isDirectory() && fs.existsSync(path.join(dir, e.name, '.git')))
            .map((e) => e.name);
    } catch {
        return null;
    }
    const preferred = children.find((c) => c.toLowerCase() === Environment.build.repoName.toLowerCase()) ?? (children.length === 1 ? children[0] : null);
    return preferred ? path.join(dir, preferred) : null;
}

/** What the panel shows: the folder that will be used and whether it can be changed there. */
export function buildDefaults() {
    return { folder: sourceFolder() ?? '', editable: Environment.geir.allowLocalPaths };
}

export function startBuild(
    session: Session,
    input: { location: string; rama: string; carpeta?: string; requerimientos?: string }
): BuildJobView {
    const branch = String(input.rama ?? '').trim().replace(/^origin\//, '');
    if (!validBranch(branch)) {
        throw new UserError('Escribe el nombre de la rama tal como está en GitLab (ej. feature/CRQ12345-sprint-1).');
    }
    const location = locate(input.location);
    const typed = String(input.carpeta ?? '').trim();
    if (typed && !Environment.geir.allowLocalPaths) {
        throw new UserError('En el servidor compartido la carpeta del repositorio se configura en el .env (GEIR_REPO_DIR).');
    }
    const folder = sourceFolder(typed);
    if (typed && !folder) {
        throw new UserError(`${typed} no es un repositorio git (no tiene .git). Escribe la carpeta donde clonaste el repositorio.`);
    }
    if (!folder && !session.credentials.password) {
        throw new UserError('Tu sesión no tiene la contraseña de red para GitLab; cierra sesión y vuelve a entrar.');
    }
    const running = [...jobs.values()].find((j) => j.owner === session.token && (j.state === 'en cola' || j.state === 'trabajando'));
    if (running) {
        return view(running);
    }
    const job: BuildJob = {
        id: crypto.randomBytes(8).toString('hex'),
        owner: session.token,
        branch,
        base: '',
        ids: [],
        state: 'en cola',
        step: 'Esperando a que termine otra compilación…',
        startedAt: new Date().toISOString(),
        finishedAt: '',
        commits: [],
        projects: [],
        files: [],
        deploy: '',
        warnings: [],
        error: '',
        log: []
    };
    jobs.set(job.id, job);
    forgetOldJobs();
    const stories = parseRequirementPairs(String(input.requerimientos ?? '')).pairs.map((p) => p.storyId);
    queue = queue.then(() => execute(session, job, location, folder, stories));
    return view(job);
}

export function buildStatus(session: Session, id: string): BuildJobView | null {
    const job = jobs.get(id);
    return job && job.owner === session.token ? view(job) : null;
}

function view(job: BuildJob): BuildJobView {
    const { owner, ...rest } = job;
    return { ...rest, log: job.log.slice(-400), commits: [...job.commits], projects: job.projects.map((p) => ({ ...p })) };
}

function forgetOldJobs() {
    const dayAgo = Date.now() - 24 * 3600 * 1000;
    for (const [id, job] of jobs) {
        if (job.finishedAt && Date.parse(job.finishedAt) < dayAgo) {
            jobs.delete(id);
        }
    }
}

async function execute(session: Session, job: BuildJob, location: DocumentLocation, local: string | null, stories: string[]) {
    const log = (line: string) => {
        job.log.push(line);
        if (job.log.length > MAX_LOG) {
            job.log.splice(0, job.log.length - MAX_LOG);
        }
    };
    const step = (text: string) => {
        job.step = text;
        log(`== ${text}`);
        console.log(`[COMPILAR ${job.branch}] ${text}`);
    };
    job.state = 'trabajando';
    try {
        const settings = buildSettings();

        // 0. The stories of this GEIR: only their commits are compiled.
        step('Leyendo las historias de la GEIR');
        job.ids = await geirIds(session, location, stories, job.warnings, log);
        if (!job.ids.length) {
            throw new UserError(
                'La GEIR no tiene tareas y no se indicaron requerimientos: no se sabe qué historias compilar. Agrega las tareas o los pares REQ ↔ historia.'
            );
        }
        log(`Historias de la GEIR: ${job.ids.join(', ')}`);

        // 1. The repository at the branch: the local folder (pull) or, only when
        //    there is none, GEIR's own clone of GitLab.
        let repo: GitRepo;
        if (local) {
            repo = new GitRepo(local, gitAuth(session), log);
            step(`Actualizando ${local} (fetch + pull de ${job.branch})`);
            await repo.fetch();
            if (!(await repo.hasRemoteBranch(job.branch))) {
                throw new UserError(`La rama ${job.branch} no existe en origin.`);
            }
            await repo.checkoutAndPull(job.branch);
        } else {
            const project = Environment.build.gitlabProject;
            const url = `${Environment.services.gitlabUrl.replace(/\/$/, '')}/${project}.git`;
            repo = new GitRepo(path.join(Environment.build.reposDir, project.replace(/[^\w.-]+/g, '_')), gitAuth(session), log);
            step(`No hay carpeta local del repositorio: descargando ${project} de GitLab (rama ${job.branch})`);
            await repo.ensureClone(url);
            await repo.fetch();
            if (!(await repo.hasRemoteBranch(job.branch))) {
                throw new UserError(`La rama ${job.branch} no existe en GitLab (${project}).`);
            }
            await repo.checkoutClean(job.branch);
        }

        // 2. The branch's commits (against the base) that name a story of the GEIR, and the
        //    files those commits changed. Commits of other stories are not compiled.
        job.base = await repo.baseBranch(settings.ramaBase);
        step(`Buscando en ${job.branch} los commits de las historias de la GEIR`);
        const all = await repo.commits(job.base, job.branch);
        const tagged = all.map((c) => ({ commit: c, related: mentionsAny(`${c.message}`, job.ids) }));
        const mine = tagged.filter((t) => t.related.length);
        for (const t of mine.filter((m) => m.commit.merge)) {
            t.commit.files = await repo.mergeFiles(t.commit.sha);
        }
        const pomList = await repo.poms();
        const view: RepoView = {
            exists: (rel) => fs.existsSync(path.join(repo.dir, rel)),
            read: (rel) => (fs.existsSync(path.join(repo.dir, rel)) ? fs.readFileSync(path.join(repo.dir, rel), 'utf8') : null),
            poms: () => pomList
        };
        // Which project each commit touches: shown next to the commit, and in the log.
        const projectNames = (files: string[]) =>
            [
                ...new Set(
                    files
                        .map((f) => projectOf(f, view, frontScript(settings.front.comando)))
                        .filter(Boolean)
                        .map((p) => (p!.kind === 'front' ? `front: ${p!.dir || '(raíz)'}` : projectLabel(view, p!)))
                )
            ].sort();
        job.commits = tagged.map(({ commit: { message, files, ...rest }, related }) => ({
            ...rest,
            related,
            projects: related.length ? projectNames(files) : []
        }));
        for (const t of mine) {
            log(
                `${t.commit.short} [${t.related.join(', ')}]${t.commit.merge ? ' (merge)' : ''}: ${t.commit.files.length} archivos → ${
                    projectNames(t.commit.files).join(', ') || 'ningún proyecto'
                }`
            );
        }
        const changed = [...new Set(mine.flatMap((t) => t.commit.files))];
        log(`${all.length} commits en la rama respecto a ${job.base}; ${mine.length} son de la GEIR (${changed.length} archivos).`);
        if (!all.length) {
            throw new UserError(`La rama ${job.branch} no tiene cambios respecto a ${job.base}: no hay nada que compilar.`);
        }
        if (!mine.length) {
            throw new UserError(
                `Ningún commit de ${job.branch} menciona las historias de la GEIR (${job.ids.join(', ')}). ` +
                    'El ID de la historia o de la tarea debe estar en el mensaje del commit.'
            );
        }
        const missing = job.ids.filter((id) => !mine.some((t) => t.related.includes(id)));
        if (missing.length) {
            job.warnings.push(`Sin commits en ${job.branch} para: ${missing.join(', ')}.`);
        }
        const plan = planBuild(changed, view, settings, installedInMaven);
        if (!plan.length) {
            throw new UserError('Los commits de la GEIR no tocan ningún proyecto con pom.xml o package.json.');
        }
        const libraries = new Set(plan.flatMap((p) => p.pom?.dependencies ?? []));
        job.projects = plan.map((p) => ({
            name: p.name,
            dir: p.dir,
            kind: p.kind,
            reason: p.reason,
            destination: p.destination,
            command: commandFor(repo.dir, p, settings, libraries),
            state: 'pendiente' as ProjectState,
            outputs: []
        }));
        log(`Proyectos a compilar: ${plan.map((p) => `${p.name} (${p.reason})`).join(', ')}`);

        // 3. Compile, dependencies first. A failure skips what depends on it.
        const failed = new Set<string>();
        const outputs: { file: string; target: string }[] = [];
        for (const [i, p] of plan.entries()) {
            const item = job.projects[i];
            const key = p.pom ? `${p.pom.groupId}:${p.pom.artifactId}` : p.dir;
            if (p.pom?.dependencies.some((d) => failed.has(d))) {
                item.state = 'omitido';
                failed.add(key);
                job.warnings.push(`${p.name} no se compiló porque falló una dependencia.`);
                continue;
            }
            item.state = 'compilando';
            step(`Compilando ${p.name} (${item.dir || 'raíz'})`);
            try {
                const found = p.kind === 'maven' ? await compileMaven(repo.dir, p, item.command, log) : await compileFront(repo.dir, p, settings, log);
                item.state = 'listo';
                if (p.destination) {
                    found.forEach((f) => outputs.push({ file: f.file, target: `${p.destination}/${f.relative}` }));
                    item.outputs = found.map((f) => `${p.destination}/${f.relative}`);
                }
            } catch (error: any) {
                item.state = 'error';
                failed.add(key);
                job.warnings.push(`${p.name}: ${firstLine(error.message)}`);
                log(String(error.message));
            }
        }

        // 4. Binaries into Deploy, next to the GEIR (the folder is created when missing).
        if (outputs.length) {
            step(`Copiando ${outputs.length} archivos a la carpeta Deploy`);
            const written = await writeDeploy(session, location, outputs, log);
            job.deploy = written.deploy;
            job.files = written.files;
            const lost = written.files.filter((f) => f.error);
            if (lost.length) {
                lost.forEach((f) => job.warnings.push(`No se subió ${f.path}: ${f.error}`));
                throw new UserError(`${lost.length} de ${written.files.length} archivos no llegaron a la carpeta Deploy. Vuelve a compilar o revisa el registro.`);
            }
        }
        if (failed.size) {
            throw new UserError(`Hubo errores al compilar: ${job.projects.filter((p) => p.state === 'error').map((p) => p.name).join(', ')}. Revisa el registro.`);
        }
        job.state = 'listo';
        step(`Listo: ${job.files.length} archivos en ${job.deploy}`);
    } catch (error: any) {
        job.state = 'error';
        job.error = String(error?.message ?? error);
        job.step = 'Terminó con error';
        log(`ERROR: ${job.error}`);
        console.error(`[COMPILAR ${job.branch}]`, job.error);
    } finally {
        job.finishedAt = new Date().toISOString();
    }
}

/**
 * IDs searched in the commit messages: the tasks of the GEIR, their parent stories in Azure
 * DevOps (commits often name the story) and the stories typed in "Requerimientos".
 */
async function geirIds(
    session: Session,
    location: DocumentLocation,
    stories: string[],
    warnings: string[],
    log: (l: string) => void
): Promise<string[]> {
    const ids = new Set<string>(stories);
    const { doc } = await loadOrCreate(session, location);
    const tasks = (await doc.listTasks()).map((t) => t.id);
    tasks.forEach((id) => ids.add(id));
    if (!MOCK && tasks.length) {
        try {
            const client = await azure(session);
            for (const id of tasks) {
                const parent = (await client.getSummary(id).catch(() => null))?.parentId;
                if (parent) {
                    ids.add(parent);
                }
            }
        } catch (error: any) {
            warnings.push(`No se pudieron leer las historias padre en Azure DevOps (${error.message}); se buscan solo los IDs de las tareas.`);
        }
    }
    log(`Tareas en la GEIR: ${tasks.join(', ') || 'ninguna'}`);
    return [...ids];
}

/** artifactId of a Maven project, or its folder name. */
function projectLabel(view: RepoView, p: { dir: string; kind: 'maven' | 'front' }): string {
    if (p.kind === 'maven') {
        const xml = view.read(p.dir ? `${p.dir}/pom.xml` : 'pom.xml');
        const artifact = xml ? /<artifactId>\s*([^<]+?)\s*<\/artifactId>/.exec(xml.replace(/<parent>[\s\S]*?<\/parent>/, ''))?.[1] : '';
        if (artifact) {
            return artifact;
        }
    }
    return p.dir ? path.posix.basename(p.dir) : '(raíz)';
}

function gitAuth(session: Session) {
    return session.credentials.password ? { username: session.names.ldap, password: session.credentials.password } : null;
}

function firstLine(text: string) {
    return String(text).split('\n')[0];
}

/** The project's own jar is already in the local Maven repository. */
function installedInMaven(pom: Pom): boolean {
    const dir = path.join(Environment.build.mavenRepo, ...pom.groupId.split('.'), pom.artifactId, pom.version);
    return fs.existsSync(path.join(dir, `${pom.artifactId}-${pom.version}.jar`));
}

function commandFor(repoDir: string, p: BuildProject, settings: BuildSettings, libraries: Set<string>): string {
    if (p.kind === 'front') {
        return settings.front.instalar ? `${settings.front.instalar} && ${settings.front.comando}` : settings.front.comando;
    }
    const isLibrary = libraries.has(`${p.pom!.groupId}:${p.pom!.artifactId}`) || !p.destination;
    let command = commandFromReadmes(repoDir, p.dir, isLibrary) ?? settings.maven.comando;
    if (isLibrary) {
        command = command.replace(/\bpackage\b/, 'install');
    }
    if (/^mvnw?(\.cmd)?\s/.test(command) && !/\s(-B|--batch-mode)\b/.test(command)) {
        command = command.replace(/^(\S+)/, '$1 -B');
    }
    if (process.env.GEIR_MAVEN_REPO && !/maven\.repo\.local/.test(command)) {
        command += ` -Dmaven.repo.local="${Environment.build.mavenRepo}"`;
    }
    return command;
}

/**
 * The README that says how to compile this project: its own, or one in a folder above it
 * (up to the repository root) with a `cd <project> && mvn …` line.
 */
function commandFromReadmes(repoDir: string, projectDir: string, mustInstall: boolean): string | null {
    let dir = projectDir;
    for (;;) {
        for (const name of ['README.md', 'readme.md', 'Readme.md', 'README.MD', 'README.txt', 'README']) {
            const file = path.join(repoDir, dir, name);
            if (fs.existsSync(file)) {
                const command = readmeCommand(fs.readFileSync(file, 'utf8'), mustInstall, dir, projectDir);
                if (command) {
                    return command;
                }
                break;
            }
        }
        if (!dir) {
            return null;
        }
        const up = path.posix.dirname(dir);
        dir = up === '.' ? '' : up;
    }
}

function toolEnv(): Record<string, string | undefined> {
    const javaHome = Environment.build.javaHome;
    if (!javaHome) {
        return {};
    }
    const key = Object.keys(process.env).find((k) => k.toLowerCase() === 'path') ?? 'PATH';
    return { JAVA_HOME: javaHome, [key]: `${path.join(javaHome, 'bin')}${path.delimiter}${process.env[key] ?? ''}` };
}

async function compileMaven(repoDir: string, p: BuildProject, command: string, log: (l: string) => void) {
    const cwd = path.join(repoDir, p.dir);
    await run(command, [], { cwd, shell: true, log, env: toolEnv(), timeoutMs: Environment.build.timeoutMinutes * 60000 });
    if (!p.destination) {
        return [];
    }
    const target = path.join(cwd, 'target');
    const files = fs.existsSync(target) ? fs.readdirSync(target) : [];
    const jar = pickJar(files, p.pom!);
    const how = { 'with-dependencies': 'assembly jar-with-dependencies', shade: 'shade (jar con dependencias)', jar: 'jar normal' }[p.pom!.packagedAs];
    if (!jar) {
        throw new Error(
            `Según su pom.xml (${how}) Maven debía dejar ${expectedJar(p.pom!)} en ${target}, pero hay: ${
                files.filter((f) => f.endsWith('.jar')).join(', ') || 'ningún .jar'
            }.`
        );
    }
    log(`Binario (${how} según el pom.xml): ${path.join(target, jar)}`);
    return [{ file: path.join(target, jar), relative: jar }];
}

async function compileFront(repoDir: string, p: BuildProject, settings: BuildSettings, log: (l: string) => void) {
    const cwd = path.join(repoDir, p.dir);
    const env: Record<string, string | undefined> = {};
    // Old webpack (md4 hashes) needs the legacy OpenSSL provider on Node 17 or newer.
    if (Number(process.versions.node.split('.')[0]) >= 17 && !/openssl-legacy-provider/.test(process.env.NODE_OPTIONS ?? '')) {
        env.NODE_OPTIONS = `${process.env.NODE_OPTIONS ?? ''} --openssl-legacy-provider`.trim();
    }
    const timeoutMs = Environment.build.timeoutMinutes * 60000;
    if (settings.front.instalar) {
        await run(settings.front.instalar, [], { cwd, shell: true, log, timeoutMs });
    }
    const started = Date.now() - 5000;
    await run(settings.front.comando, [], { cwd, shell: true, log, env, timeoutMs });
    let out = path.join(cwd, settings.front.salida);
    if (!fs.existsSync(out)) {
        // Output elsewhere (webpack output.path): the folder of the bundle.js just built.
        const bundle = findNewest(cwd, (name) => name === 'bundle.js', started);
        if (!bundle) {
            throw new Error(
                `"${settings.front.comando}" terminó pero no existe ${out} ni se generó un bundle.js. Indica la carpeta en config → compilacion.front.salida.`
            );
        }
        log(`No existe ${out}; se usa la carpeta del bundle.js generado: ${path.dirname(bundle)}`);
        out = path.dirname(bundle);
    }
    const files: { file: string; relative: string }[] = [];
    const walk = (dir: string): void => {
        for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
            const full = path.join(dir, e.name);
            if (e.isDirectory()) {
                walk(full);
            } else {
                files.push({ file: full, relative: path.relative(out, full).split(path.sep).join('/') });
            }
        }
    };
    walk(out);
    log(`Front end: ${files.length} archivos en ${out}`);
    return files;
}

/** Newest file (modified after `since`) whose name matches, outside node_modules and .git. */
function findNewest(dir: string, match: (name: string) => boolean, since: number): string | null {
    let best: { file: string; time: number } | null = null;
    const walk = (current: string, depth: number): void => {
        if (depth > 8) {
            return;
        }
        for (const e of fs.readdirSync(current, { withFileTypes: true })) {
            if (e.name === 'node_modules' || e.name === '.git') {
                continue;
            }
            const full = path.join(current, e.name);
            if (e.isDirectory()) {
                walk(full, depth + 1);
            } else if (match(e.name)) {
                const time = fs.statSync(full).mtimeMs;
                if (time >= since && (!best || time > best.time)) {
                    best = { file: full, time };
                }
            }
        }
    };
    walk(dir, 0);
    return (best as { file: string } | null)?.file ?? null;
}

async function writeDeploy(
    session: Session,
    location: DocumentLocation,
    outputs: { file: string; target: string }[],
    log: (l: string) => void
): Promise<{ deploy: string; files: { path: string; size: number; error?: string }[] }> {
    const files: { path: string; size: number; error?: string }[] = [];
    if (location.kind === 'local') {
        const root = deployRoot(path.dirname(location.filePath));
        for (const o of outputs) {
            const dest = path.join(root, ...o.target.split('/'));
            fs.mkdirSync(path.dirname(dest), { recursive: true });
            fs.copyFileSync(o.file, dest);
            files.push({ path: `Deploy/${o.target}`, size: fs.statSync(dest).size });
            log(`→ ${dest}`);
        }
        return { deploy: root, files };
    }
    return withMicrosoft(session, async (client) => {
        const { siteUrl, serverRelativePath } = await client.normalize(location);
        const root = deployRoot(path.posix.dirname(serverRelativePath));
        // One file failing (proxy, size, lock) does not stop the others; each one is checked
        // in OneDrive afterwards (same size) and the page lists what did not arrive.
        for (const o of outputs) {
            const data = fs.readFileSync(o.file);
            const target = `${root}/${o.target}`;
            try {
                const started = Date.now();
                await client.write({ kind: 'sharepoint', siteUrl, serverRelativePath: target }, data);
                const size = await client.size({ kind: 'sharepoint', siteUrl, serverRelativePath: target });
                if (size !== data.length) {
                    throw new Error(`en OneDrive quedó con ${size ?? 0} bytes de ${data.length}`);
                }
                files.push({ path: `Deploy/${o.target}`, size: data.length });
                log(`→ ${target} (${Math.round(data.length / 1024)} KB, ${((Date.now() - started) / 1000).toFixed(1)} s)`);
            } catch (error: any) {
                if (error instanceof SessionExpiredError) {
                    throw error;
                }
                files.push({ path: `Deploy/${o.target}`, size: data.length, error: String(error?.message ?? error) });
                log(`ERROR al subir ${target}: ${error?.message ?? error}`);
            }
        }
        return { deploy: describeLocation({ kind: 'sharepoint', siteUrl, serverRelativePath: root }), files };
    });
}
