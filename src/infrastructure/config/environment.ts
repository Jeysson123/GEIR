import * as dotenv from 'dotenv';
import * as os from 'os';
import * as path from 'path';

// .env first, then .env.local (git-ignored): the services of this installation (server
// names, domain, project) live there, never in the repository.
dotenv.config();
dotenv.config({ path: path.resolve(process.cwd(), '.env.local') });

/** Everything configurable through .env (see .env.example). No credentials here. */
export const Environment = {
    // Services GEIR reads from.
    services: {
        azureDevOpsUrl: process.env.AZURE_DEVOPS_URL ?? '',
        gitlabUrl: process.env.GITLAB_BASE_URL ?? '',
        timeout: 30000
    },

    // GEIR web app (npm start)
    geir: {
        port: Number(process.env.GEIR_PORT ?? 4300),
        // 127.0.0.1: only this machine. On a shared server: 0.0.0.0 (behind an HTTPS proxy).
        host: process.env.GEIR_HOST ?? '127.0.0.1',
        // Behind a reverse proxy (nginx / IIS) that terminates HTTPS.
        trustProxy: process.env.GEIR_TRUST_PROXY === 'true',
        // Paths on the GEIR machine itself ("C:\\...\\N - GEIR.xlsx"). Off on a shared server:
        // users would read and write the server's disk. Default: only when bound to localhost.
        allowLocalPaths:
            process.env.GEIR_ALLOW_LOCAL_PATHS !== undefined
                ? process.env.GEIR_ALLOW_LOCAL_PATHS === 'true'
                : ['127.0.0.1', 'localhost', '::1'].includes(process.env.GEIR_HOST ?? '127.0.0.1'),
        // Browser for background work: Edge on Windows, Playwright's Chromium on Linux.
        browserChannel: process.env.GEIR_BROWSER ?? (process.platform === 'win32' ? 'msedge' : 'chromium'),
        // Use fixture data instead of Azure DevOps / GitLab (demo or off-VPN work).
        mock: process.env.GEIR_MOCK === 'true' || process.argv.includes('--mock'),
        templatePath: process.env.GEIR_TEMPLATE ?? 'templates/N - GEIR.xlsx',
        fileName: 'N - GEIR.xlsx',
        azure: {
            collection: process.env.AZURE_COLLECTION ?? '',
            project: process.env.AZURE_PROJECT ?? '',
            team: process.env.AZURE_TEAM ?? ''
        },
        gitlab: {
            // Projects searched for commits / merge requests that mention the task ID.
            projects: (process.env.GITLAB_PROJECTS ?? '')
                .split(',')
                .map((p) => p.trim())
                .filter(Boolean)
        },
        microsoft: {
            emailDomain: process.env.MICROSOFT_EMAIL_DOMAIN ?? ''
        },
        sharepointHost: process.env.SHAREPOINT_SITE ?? ''
    },

    // Compilation of the binaries (admin option "Rama"): git, Maven and npm of this machine.
    build: {
        // Local clone of the source repository. Empty: GEIR looks for GEIR_REPO_NAME next to itself,
        // then keeps its own clone of GitLab.
        repoName: process.env.GEIR_REPO_NAME ?? '',
        localRepo: process.env.GEIR_REPO_DIR ?? '',
        // GitLab project cloned when there is no local folder.
        gitlabProject: process.env.GEIR_REPO_PROJECT ?? (process.env.GITLAB_PROJECTS ?? '').split(',')[0].trim(),
        // Where GEIR keeps its clone (shared by every user; one build at a time).
        reposDir: process.env.GEIR_REPOS_DIR ?? path.join(os.tmpdir(), 'geir-repos'),
        git: process.env.GEIR_GIT ?? 'git',
        // Self-hosted GitLab often uses an internal certificate (the rest of GEIR ignores it too).
        gitSslVerify: process.env.GEIR_GIT_SSL_VERIFY === 'true',
        // JDK for Maven. Empty: the system one.
        javaHome: process.env.GEIR_JAVA_HOME ?? '',
        // Maven local repository, if not ~/.m2/repository.
        mavenRepo: process.env.GEIR_MAVEN_REPO ?? path.join(os.homedir(), '.m2', 'repository'),
        timeoutMinutes: Number(process.env.GEIR_BUILD_TIMEOUT_MIN ?? 30)
    }

} as const;