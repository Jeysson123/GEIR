import * as fs from 'fs';
import * as path from 'path';
import { Environment } from '../config/environment';
import { LogLine, run } from './shell';

export interface Commit {
    sha: string;
    short: string;
    author: string;
    date: string;
    title: string;
    message: string;
    files: string[];
    /** Merge commit (e.g. a merge request): its files are what it brought into the branch. */
    merge: boolean;
}

export interface GitAuth {
    username: string;
    password: string;
}

// Credentials go to git through environment variables of the child process only: an inline
// credential helper prints them when git asks. Nothing is written to disk or to .git/config.
const HELPER =
    '!f() { test "$1" = get && printf "username=%s\\npassword=%s\\n" "$GEIR_GIT_USER" "$GEIR_GIT_PASS"; }; f';

/** A clone of the source repository: the local folder or GEIR's own copy of GitLab. */
export class GitRepo {
    constructor(readonly dir: string, private auth: GitAuth | null, private log: LogLine) {}

    git(args: string[], cwd = this.dir, echo = true): Promise<string> {
        // core.longpaths: Java projects can have paths over Windows' 260-character limit.
        const config = ['-c', 'core.quotepath=off', '-c', 'core.longpaths=true'];
        if (this.auth) {
            config.push('-c', 'credential.helper=', '-c', `credential.helper=${HELPER}`);
        }
        if (!Environment.build.gitSslVerify) {
            config.push('-c', 'http.sslVerify=false');
        }
        return run(Environment.build.git, [...config, ...args], {
            cwd,
            log: this.log,
            env: {
                GIT_TERMINAL_PROMPT: '0',
                GEIR_GIT_USER: this.auth?.username,
                GEIR_GIT_PASS: this.auth?.password
            },
            secrets: this.auth ? [this.auth.password, encodeURIComponent(this.auth.password)] : [],
            timeoutMs: 20 * 60 * 1000,
            display: `git ${args.join(' ')}`,
            echo
        });
    }

    private async quiet(args: string[]): Promise<string | null> {
        try {
            return (await this.git(args, this.dir, false)).trim();
        } catch {
            return null;
        }
    }

    /** Clones the repository the first time (GitLab mode). */
    async ensureClone(url: string) {
        if (fs.existsSync(path.join(this.dir, '.git'))) {
            await this.git(['remote', 'set-url', 'origin', url]);
            return;
        }
        fs.mkdirSync(path.dirname(this.dir), { recursive: true });
        await this.git(['clone', '--no-checkout', url, this.dir], path.dirname(this.dir));
        await this.git(['config', 'core.longpaths', 'true']);
    }

    fetch() {
        return this.git(['fetch', 'origin', '--prune']);
    }

    async hasRemoteBranch(branch: string): Promise<boolean> {
        return (await this.quiet(['rev-parse', '--verify', '--quiet', `refs/remotes/origin/${branch}`])) !== null;
    }

    /** GEIR's own clone: exactly what is in GitLab, discarding anything left by an earlier build. */
    async checkoutClean(branch: string) {
        await this.git(['checkout', '-f', '-B', branch, `origin/${branch}`]);
        await this.git(['clean', '-fd']);
    }

    /** The user's own folder: refuses to touch uncommitted work, then checkout + pull. */
    async checkoutAndPull(branch: string) {
        const status = await this.git(['status', '--porcelain', '--untracked-files=no'], this.dir, false);
        const changed = status
            .split('\n')
            .filter((l) => l.trim())
            .map((l) => l.slice(3).trim());
        // `npm install` of an earlier build may have rewritten the lock file: that one is ours.
        const locks = changed.filter((f) => /(^|\/)(package-lock|npm-shrinkwrap)\.json$/.test(f));
        const others = changed.filter((f) => !locks.includes(f));
        if (others.length) {
            throw new Error(
                `La carpeta ${this.dir} tiene cambios sin confirmar (${others.slice(0, 3).join(', ')}${others.length > 3 ? '…' : ''}); ` +
                    'GEIR no hace checkout encima. Haz commit o stash y vuelve a compilar.'
            );
        }
        if (locks.length) {
            await this.git(['checkout', '--', ...locks]);
        }
        await this.git(['checkout', branch]);
        await this.git(['pull', '--ff-only', 'origin', branch]);
    }

    /** Base the branch is compared with: the configured one or the remote's default branch. */
    async baseBranch(configured: string): Promise<string> {
        const candidates = configured ? [configured] : [];
        const head = await this.quiet(['symbolic-ref', '--short', 'refs/remotes/origin/HEAD']);
        if (head) {
            candidates.push(head.replace(/^origin\//, ''));
        }
        candidates.push('master', 'main', 'develop');
        for (const name of candidates) {
            if (await this.hasRemoteBranch(name)) {
                return name;
            }
        }
        throw new Error('No se encontró la rama base (master/main/develop). Indícala en config/geir.config.json → compilacion.ramaBase.');
    }

    /**
     * Commits of the branch that are not in the base, with their full message and files.
     * Merges are included: the story ID is often only in the merge of the merge request
     * ("Merge branch 'feature/100008-report'…"). Their files come from `mergeFiles()`.
     */
    async commits(base: string, branch: string): Promise<Commit[]> {
        const text = await this.git(
            [
                'log',
                '--date=short',
                '--name-only',
                '--format=%x1e%H%x1f%h%x1f%an%x1f%ad%x1f%P%x1f%s%x1f%B%x1d',
                `origin/${base}..origin/${branch}`
            ],
            this.dir,
            false
        );
        return text
            .split('\x1e')
            .filter((chunk) => chunk.includes('\x1f'))
            .map((chunk) => {
                const [head, names = ''] = chunk.split('\x1d');
                const [sha, short, author, date, parents, title, message] = head.split('\x1f');
                const files = names
                    .split('\n')
                    .map((l) => l.trim())
                    .filter(Boolean);
                const merge = parents.trim().split(/\s+/).length > 1;
                return { sha, short, author, date, title, message: (message ?? title).trim(), files, merge };
            });
    }

    /** Files a merge brought into the branch (against its first parent). */
    async mergeFiles(sha: string): Promise<string[]> {
        const text = await this.git(['diff', '--name-only', `${sha}^1`, sha], this.dir, false);
        return text.split('\n').map((l) => l.trim()).filter(Boolean);
    }

    /** Files the branch changed since it left the base. */
    async changedFiles(base: string, branch: string): Promise<string[]> {
        const text = await this.git(['diff', '--name-only', `origin/${base}...origin/${branch}`], this.dir, false);
        return text.split('\n').map((l) => l.trim()).filter(Boolean);
    }

    /** Every pom.xml tracked in the branch. */
    async poms(): Promise<string[]> {
        const text = await this.git(['ls-files', '--', 'pom.xml', '*/pom.xml'], this.dir, false);
        return text.split('\n').map((l) => l.trim()).filter(Boolean);
    }
}
