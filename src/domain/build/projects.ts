import * as path from 'path';

/**
 * Which projects of the source repository a branch touched, in what order they compile
 * and where each binary goes in the Deploy folder. Pure functions: the repository is read
 * through the `RepoView` given by the caller.
 */

export interface RepoView {
    exists(rel: string): boolean;
    read(rel: string): string | null;
    /** Every pom.xml of the repository (relative paths), to find who depends on whom. */
    poms(): string[];
}

export interface Pom {
    groupId: string;
    artifactId: string;
    version: string;
    packaging: string;
    dependencies: string[]; // "groupId:artifactId"
    /**
     * What `package` leaves in target/, read from <build>: the assembly plugin with the
     * jar-with-dependencies descriptor (…-jar-with-dependencies.jar), the shade plugin (the
     * main jar already carries the dependencies) or a plain jar.
     */
    packagedAs: 'with-dependencies' | 'shade' | 'jar';
    /** <build><finalName>, resolved; default artifactId-version. */
    finalName: string;
}

export interface BuildRule {
    /** Regular expression on the artifactId (Maven) or the folder name (front end). */
    coincide: string;
    /** Folder inside Deploy; empty: only installed in the local Maven repository. */
    destino: string;
}

export interface BuildSettings {
    ramaBase: string;
    /**
     * true: when a library of the repository changes (a shared core), also rebuild every
     * project that uses it. false (default): only the projects the GEIR's commits touch.
     */
    recompilarDependientes: boolean;
    maven: { comando: string; reglas: BuildRule[] };
    front: { instalar: string; comando: string; salida: string; destino: string };
}

export interface BuildProject {
    dir: string; // relative to the repository, '' for the root
    kind: 'maven' | 'front';
    name: string; // artifactId or folder name
    pom?: Pom;
    reason: 'cambio' | 'depende' | 'falta';
    /** Folder inside Deploy, '' when nothing is copied (shared libraries). */
    destination: string;
}

export const DEFAULT_BUILD: BuildSettings = {
    ramaBase: '',
    recompilarDependientes: false,
    maven: {
        comando: 'mvn clean install -DskipTests',
        reglas: [
            { coincide: '-core$', destino: '' },
            { coincide: 'interfaces', destino: 'backend/jar_interfaces' },
            { coincide: '.*', destino: 'backend' }
        ]
    },
    front: {
        instalar: '', // the front end is built with the node_modules it already has
        comando: 'npm run buildprod',
        salida: 'dist/js',
        destino: 'frontend/dist/js'
    }
};

/** The commit message names one of the IDs (as a whole number: 2826 does not match 28261). */
export function mentionsAny(text: string, ids: Iterable<string>): string[] {
    const found: string[] = [];
    for (const id of ids) {
        if (id && new RegExp(`(^|\\D)${id}(\\D|$)`).test(text ?? '')) {
            found.push(id);
        }
    }
    return found;
}

const SAFE_BRANCH = /^(?!-)(?!.*\.\.)(?!.*\/\/)[A-Za-z0-9._\/-]{1,200}$/;

/** Branch names go to git; only the characters branch names really use. */
export function validBranch(name: string): boolean {
    return SAFE_BRANCH.test(name) && !name.endsWith('/') && !name.endsWith('.lock');
}

/** Reads what the build needs from a pom.xml (inherits groupId/version from <parent>). */
export function parsePom(xml: string): Pom {
    const clean = xml.replace(/<!--[\s\S]*?-->/g, '');
    const parent = /<parent>([\s\S]*?)<\/parent>/.exec(clean)?.[1] ?? '';
    // Own coordinates: the tags of <project> outside <parent>, <dependencies>, <build>...
    const own = clean
        .replace(/<parent>[\s\S]*?<\/parent>/, '')
        .replace(/<(dependencies|dependencyManagement|build|profiles|plugins|reporting|repositories|pluginRepositories|properties|prerequisites)>[\s\S]*?<\/\1>/g, '');
    const tag = (text: string, name: string) => new RegExp(`<${name}>\\s*([^<]+?)\\s*</${name}>`).exec(text)?.[1] ?? '';
    const deps = [...(/<dependencies>([\s\S]*?)<\/dependencies>/.exec(
        clean.replace(/<dependencyManagement>[\s\S]*?<\/dependencyManagement>/, '').replace(/<build>[\s\S]*?<\/build>/, '')
    )?.[1] ?? '').matchAll(/<dependency>([\s\S]*?)<\/dependency>/g)].map((m) => `${tag(m[1], 'groupId')}:${tag(m[1], 'artifactId')}`);
    const groupId = tag(own, 'groupId') || tag(parent, 'groupId');
    const artifactId = tag(own, 'artifactId');
    const version = tag(own, 'version') || tag(parent, 'version');

    // Only the plugins that run: <build><plugins>, not <pluginManagement> or <profiles>.
    const build = (/<build>([\s\S]*?)<\/build>/.exec(clean.replace(/<profiles>[\s\S]*?<\/profiles>/, ''))?.[1] ?? '').replace(
        /<pluginManagement>[\s\S]*?<\/pluginManagement>/,
        ''
    );
    const plugins = [...build.matchAll(/<plugin>([\s\S]*?)<\/plugin>/g)].map((m) => m[1]);
    const assembly = plugins.find(
        (p) => tag(p, 'artifactId') === 'maven-assembly-plugin' && /<descriptorRef>\s*jar-with-dependencies\s*<\/descriptorRef>/.test(p)
    );
    const shade = plugins.find((p) => tag(p, 'artifactId') === 'maven-shade-plugin');
    const appendId = !assembly || !/<appendAssemblyId>\s*false\s*<\/appendAssemblyId>/.test(assembly);

    const properties: Record<string, string> = {};
    for (const m of (/<properties>([\s\S]*?)<\/properties>/.exec(clean)?.[1] ?? '').matchAll(/<([\w.-]+)>\s*([^<]*?)\s*<\/\1>/g)) {
        properties[m[1]] = m[2];
    }
    const resolve = (text: string) =>
        text.replace(/\$\{([^}]+)\}/g, (all, name: string) =>
            name === 'project.artifactId' || name === 'artifactId'
                ? artifactId
                : name === 'project.version' || name === 'version'
                ? version
                : name === 'project.groupId'
                ? groupId
                : properties[name] ?? all
        );
    const finalName = resolve(tag(build.replace(/<plugins>[\s\S]*?<\/plugins>/, ''), 'finalName')) || `${artifactId}-${version}`;
    return {
        groupId,
        artifactId,
        version,
        packaging: tag(own, 'packaging') || 'jar',
        dependencies: deps,
        packagedAs: assembly && appendId ? 'with-dependencies' : assembly || shade ? 'shade' : 'jar',
        finalName
    };
}

/** "npm run buildprod" → "buildprod". */
export function frontScript(command: string): string {
    return /\bnpm\s+run(?:-script)?\s+([\w:.-]+)/.exec(command)?.[1] ?? 'buildprod';
}

/**
 * Nearest folder (from the file up to the root) with a pom.xml or a package.json. A
 * package.json that has the front end's build script (buildprod) is the front end even when
 * the folder also has a pom.xml.
 */
export function projectOf(file: string, repo: RepoView, script = 'buildprod'): { dir: string; kind: 'maven' | 'front' } | null {
    let dir = path.posix.dirname(file.replace(/\\/g, '/'));
    for (;;) {
        const at = (name: string) => (dir === '.' ? name : `${dir}/${name}`);
        if (!/(^|\/)node_modules(\/|$)/.test(dir) && repo.exists(at('package.json'))) {
            const scripts = (() => {
                try {
                    return JSON.parse(repo.read(at('package.json')) ?? '{}').scripts ?? {};
                } catch {
                    return {};
                }
            })();
            if (scripts[script]) {
                return { dir: dir === '.' ? '' : dir, kind: 'front' };
            }
        }
        if (repo.exists(at('pom.xml'))) {
            return { dir: dir === '.' ? '' : dir, kind: 'maven' };
        }
        if (dir === '.' || dir === '' || dir === '/') {
            return null;
        }
        dir = path.posix.dirname(dir);
    }
}

function destinationFor(name: string, rules: BuildRule[]): string {
    for (const rule of rules) {
        if (new RegExp(rule.coincide, 'i').test(name)) {
            return rule.destino;
        }
    }
    return '';
}

/**
 * The projects to compile, dependencies first: the ones the given files belong to (the files
 * of the GEIR's commits), optionally the projects that depend on them, and the repository's
 * own libraries they need that are missing from the local Maven repository.
 */
export function planBuild(
    changedFiles: string[],
    repo: RepoView,
    settings: BuildSettings,
    installed: (pom: Pom) => boolean
): BuildProject[] {
    const poms = new Map<string, { dir: string; pom: Pom }>(); // "g:a" → project
    for (const rel of repo.poms()) {
        const xml = repo.read(rel);
        if (!xml) {
            continue;
        }
        const pom = parsePom(xml);
        if (pom.artifactId && pom.packaging !== 'pom') {
            const dir = path.posix.dirname(rel);
            poms.set(`${pom.groupId}:${pom.artifactId}`, { dir: dir === '.' ? '' : dir, pom });
        }
    }

    const chosen = new Map<string, BuildProject>(); // dir → project
    const addMaven = (key: string, reason: BuildProject['reason']) => {
        const p = poms.get(key);
        if (p && !chosen.has(p.dir)) {
            chosen.set(p.dir, {
                dir: p.dir,
                kind: 'maven',
                name: p.pom.artifactId,
                pom: p.pom,
                reason,
                destination: destinationFor(p.pom.artifactId, settings.maven.reglas)
            });
        }
    };

    for (const file of changedFiles) {
        const found = projectOf(file, repo, frontScript(settings.front.comando));
        if (!found || chosen.has(found.dir)) {
            continue;
        }
        if (found.kind === 'front') {
            const name = found.dir ? path.posix.basename(found.dir) : 'front';
            chosen.set(found.dir, { dir: found.dir, kind: 'front', name, reason: 'cambio', destination: settings.front.destino });
            continue;
        }
        const xml = repo.read(found.dir ? `${found.dir}/pom.xml` : 'pom.xml');
        const pom = xml ? parsePom(xml) : null;
        if (pom?.packaging === 'pom') {
            // Parent/aggregator pom changed: every module under it.
            for (const [key, p] of poms) {
                if (!found.dir || p.dir.startsWith(`${found.dir}/`)) {
                    addMaven(key, 'cambio');
                }
            }
        } else if (pom) {
            addMaven(`${pom.groupId}:${pom.artifactId}`, 'cambio');
        }
    }

    // Dependents of what changed (transitively), only when configured: their fat jars embed
    // the changed code, but the GEIR asks for the projects its stories touched.
    let grew = settings.recompilarDependientes;
    while (grew) {
        grew = false;
        const built = new Set([...chosen.values()].filter((p) => p.pom).map((p) => `${p.pom!.groupId}:${p.pom!.artifactId}`));
        for (const [key, p] of poms) {
            if (!chosen.has(p.dir) && p.pom.dependencies.some((d) => built.has(d))) {
                addMaven(key, 'depende');
                grew = true;
            }
        }
    }

    // Libraries of the repository the chosen projects need and the server does not have yet.
    grew = true;
    while (grew) {
        grew = false;
        for (const p of [...chosen.values()]) {
            for (const dep of p.pom?.dependencies ?? []) {
                const lib = poms.get(dep);
                if (lib && !chosen.has(lib.dir) && !installed(lib.pom)) {
                    addMaven(dep, 'falta');
                    grew = true;
                }
            }
        }
    }

    return order([...chosen.values()]);
}

/** Dependencies before the projects that use them; front ends last. */
function order(projects: BuildProject[]): BuildProject[] {
    const byKey = new Map(projects.filter((p) => p.pom).map((p) => [`${p.pom!.groupId}:${p.pom!.artifactId}`, p]));
    const done = new Set<BuildProject>();
    const out: BuildProject[] = [];
    const visit = (p: BuildProject, trail: Set<BuildProject>) => {
        if (done.has(p) || trail.has(p)) {
            return;
        }
        trail.add(p);
        for (const dep of p.pom?.dependencies ?? []) {
            const d = byKey.get(dep);
            if (d) {
                visit(d, trail);
            }
        }
        done.add(p);
        out.push(p);
    };
    projects
        .filter((p) => p.kind === 'maven')
        .sort((a, b) => a.dir.localeCompare(b.dir))
        .forEach((p) => visit(p, new Set()));
    return [...out, ...projects.filter((p) => p.kind === 'front')];
}

const MAVEN = /^mvnw?(\.cmd)?\s/;
const BUILDS = /\b(package|install)\b/;
const UNSAFE = /[;&|<>`$]/;

/**
 * Build command of a project taken from a README, if any. `readmeDir` is where the README
 * is and `projectDir` the project (both relative to the repository, '' for the root):
 *   - the project's own README: its first `mvn … package|install` line;
 *   - a README higher up: only a line aimed at this project, `cd <project> && mvn …`
 *     (a bare `mvn` there would compile every module, not just the GEIR's ones).
 * Libraries other projects use are always installed so they compile against the new code.
 */
export function readmeCommand(readme: string | null, mustInstall: boolean, readmeDir = '', projectDir = ''): string | null {
    if (!readme) {
        return null;
    }
    const norm = (p: string) => p.replace(/\\/g, '/').replace(/^\.\/|\/+$/g, '').replace(/^"|"$/g, '');
    const own = norm(readmeDir) === norm(projectDir);
    for (const raw of readme.split(/\r?\n/)) {
        const line = raw.replace(/^\s*(?:[$>]\s*|`+|\d+\.\s+|[-*]\s+)+/, '').replace(/`+\s*$/, '').trim();
        const cd = /^cd\s+("[^"]+"|\S+)\s*&&\s*(.+)$/.exec(line);
        const command = cd ? cd[2].trim() : line;
        if (!MAVEN.test(command) || !BUILDS.test(command) || UNSAFE.test(command)) {
            continue;
        }
        if (cd) {
            const target = norm(path.posix.join(norm(readmeDir) || '.', norm(cd[1])));
            if (target !== norm(projectDir)) {
                continue;
            }
        } else if (!own) {
            continue;
        }
        return mustInstall ? command.replace(/\bpackage\b/, 'install') : command;
    }
    return null;
}

/** The jar `package` leaves in target/ according to the pom.xml. */
export function expectedJar(pom: Pom): string {
    return pom.packagedAs === 'with-dependencies' ? `${pom.finalName}-jar-with-dependencies.jar` : `${pom.finalName}.jar`;
}

/**
 * The jar to deploy, as the pom.xml says: with the assembly plugin the
 * …-jar-with-dependencies.jar; otherwise (shade or plain) the main jar. Never original-…,
 * -sources or -javadoc. Null when Maven did not leave it.
 */
export function pickJar(files: string[], pom: Pom): string | null {
    const wanted = expectedJar(pom);
    return files.find((f) => f === wanted) ?? files.find((f) => f.toLowerCase() === wanted.toLowerCase()) ?? null;
}

/**
 * Where the Deploy folder goes: next to the GEIR, or next to its "GEIR" folder when the
 * document lives in one (Sprint/GEIR/N - GEIR.xlsx → Sprint/Deploy).
 */
export function deployRoot(geirFolder: string): string {
    const clean = geirFolder.replace(/[\\/]+$/, '');
    const parent = path.posix.dirname(clean.replace(/\\/g, '/'));
    const base = path.posix.basename(clean.replace(/\\/g, '/'));
    const root = base.toLowerCase() === 'geir' && parent && parent !== '.' ? clean.slice(0, clean.length - base.length - 1) : clean;
    return `${root}${clean.includes('\\') && !clean.includes('/') ? '\\' : '/'}Deploy`;
}
