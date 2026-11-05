import { Browser, BrowserContext, Page, chromium } from '@playwright/test';
import * as crypto from 'crypto';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { Environment } from '../infrastructure/config/environment';
import type { MfaPrompt } from '../infrastructure/microsoft/microsoft-login';

// Several people share the server: no sign-in (cookies, tokens) is ever written to disk.
// Each session keeps its own in memory, and the user's browser keeps the Microsoft token in
// sessionStorage so a new login in the same tab does not ask for the code again.

const IDLE_TIMEOUT_MS = 8 * 60 * 60 * 1000;
const WORK_ROOT = path.join(os.tmpdir(), 'geir');

export interface NetworkCredentials {
    username: string; // network user: Azure DevOps (Windows auth) and GitLab (LDAP)
    password: string;
    email?: string; // Microsoft 365 account (OneDrive)
    microsoftRefreshToken?: string; // kept by this user's browser (sessionStorage), if any
}

/** Account names derived from the single network login the user types. */
export function accountNames(username: string, email?: string) {
    const trimmed = username.trim();
    const bare = trimmed.includes('\\') ? trimmed.split('\\').pop()! : trimmed.split('@')[0];
    return {
        // Azure DevOps Server (Windows auth) accepts DOMAIN\user or user@domain as typed.
        windows: trimmed,
        // GitLab LDAP sign-in uses the bare account name.
        ldap: bare,
        // Microsoft 365 needs the e-mail address (typed at login; the domain is a demo fallback).
        email: email?.trim()
            ? email.trim()
            : trimmed.includes('@')
            ? trimmed
            : `${bare}@${Environment.geir.microsoft.emailDomain}`
    };
}

let sharedBrowser: Promise<Browser> | null = null;

/**
 * One headless browser for background work, launched on first use. Every session uses its
 * own contexts in it (separate cookies), never a shared one.
 */
export function browser(): Promise<Browser> {
    if (!sharedBrowser) {
        const channel = Environment.geir.browserChannel;
        const options = channel && channel !== 'chromium' ? { channel, headless: true } : { headless: true };
        const launching: Promise<Browser> = chromium.launch(options).then(
            (launched) => {
                // Edge may be closed from outside (managed VM policies): launch a new one next time.
                launched.on('disconnected', () => {
                    if (sharedBrowser === launching) {
                        console.log('[GEIR] Edge se cerró; se abrirá uno nuevo en la próxima operación.');
                        sharedBrowser = null;
                    }
                });
                return launched;
            },
            (error) => {
                sharedBrowser = null;
                throw error;
            }
        );
        sharedBrowser = launching;
    }
    return sharedBrowser;
}

export async function closeBrowser() {
    const pending = sharedBrowser;
    sharedBrowser = null;
    await (await pending?.catch(() => null))?.close();
}

/**
 * A signed-in user. Credentials live only in this process's memory; the browser only
 * gets a random token (kept in sessionStorage).
 */
export class Session {
    readonly token = crypto.randomBytes(24).toString('hex');
    lastSeen = Date.now();
    displayName: string;
    // Microsoft MFA in progress, shown by the UI; and a code typed there for it.
    mfaPrompt: MfaPrompt | null = null;
    mfaCode: string | null = null;
    // Files made by the admin options, kept to download from the page (name → content).
    adminFiles = new Map<string, Buffer>();
    // Page of the Microsoft sign-in in progress, so the UI can show what it displays.
    microsoftPage: Page | null = null;
    // Microsoft refresh token of this user (memory only). `microsoftTokenChanged` tells the
    // server to send it to the browser, which keeps it in sessionStorage.
    microsoftRefreshToken: string | null;
    microsoftTokenChanged = false;
    // Temporary files of this session (OneDrive downloads), removed when it ends.
    readonly workDir = path.join(WORK_ROOT, crypto.randomBytes(8).toString('hex'));
    private contexts = new Map<string, Promise<BrowserContext>>();
    private closers: (() => Promise<void> | void)[] = [];

    constructor(readonly credentials: NetworkCredentials) {
        this.displayName = credentials.username;
        this.microsoftRefreshToken = credentials.microsoftRefreshToken || null;
        delete this.credentials.microsoftRefreshToken;
    }

    setMicrosoftRefreshToken(token: string | null) {
        this.microsoftRefreshToken = token;
        this.microsoftTokenChanged = true;
    }

    /** Runs when the session ends (logout or idle timeout). */
    onClose(fn: () => Promise<void> | void) {
        this.closers.push(fn);
    }

    get names() {
        return accountNames(this.credentials.username, this.credentials.email);
    }

    /** Creates (once) and returns a named browser context. */
    context(key: string, factory: () => Promise<BrowserContext>): Promise<BrowserContext> {
        let ctx = this.contexts.get(key);
        if (!ctx) {
            const created: Promise<BrowserContext> = factory().then(
                (c) => {
                    // Forget it once closed (e.g. Edge closed from outside): recreate next time.
                    c.on('close', () => {
                        if (this.contexts.get(key) === created) {
                            this.contexts.delete(key);
                        }
                    });
                    return c;
                },
                (error) => {
                    this.contexts.delete(key);
                    throw error;
                }
            );
            ctx = created;
            this.contexts.set(key, ctx);
        }
        return ctx;
    }

    resetContext(key: string) {
        const ctx = this.contexts.get(key);
        this.contexts.delete(key);
        ctx?.then((c) => c.close()).catch(() => undefined);
    }

    async close() {
        const all = [...this.contexts.values()];
        this.contexts.clear();
        await Promise.all(all.map((p) => p.then((c) => c.close()).catch(() => undefined)));
        await Promise.all(this.closers.splice(0).map((fn) => Promise.resolve().then(fn).catch(() => undefined)));
        this.adminFiles.clear();
        this.microsoftRefreshToken = null;
        this.credentials.password = '';
        fs.rmSync(this.workDir, { recursive: true, force: true });
    }
}

const sessions = new Map<string, Session>();

export function createSession(credentials: NetworkCredentials): Session {
    const session = new Session(credentials);
    sessions.set(session.token, session);
    return session;
}

/** Ends sessions idle for too long, so passwords and tokens do not stay in memory. */
setInterval(() => {
    for (const [token, session] of sessions) {
        if (Date.now() - session.lastSeen > IDLE_TIMEOUT_MS) {
            void destroySession(token);
        }
    }
}, 15 * 60 * 1000).unref();

export function getSession(token: string | undefined): Session | null {
    if (!token) {
        return null;
    }
    const session = sessions.get(token);
    if (!session) {
        return null;
    }
    if (Date.now() - session.lastSeen > IDLE_TIMEOUT_MS) {
        void destroySession(token);
        return null;
    }
    session.lastSeen = Date.now();
    return session;
}

export async function destroySession(token: string) {
    const session = sessions.get(token);
    sessions.delete(token);
    await session?.close();
}
