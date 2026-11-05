import { Page } from '@playwright/test';

/** What Microsoft is asking the user to do for MFA, so a UI can show it. */
export interface MfaPrompt {
    kind: 'approve' | 'code' | 'device';
    number: string; // number to match in Authenticator, or the device code to type
    message: string;
    url?: string; // device sign-in: page where the code is typed
    // Device sign-in done by GEIR itself: 'working', 'approve' (approveNumber in
    // Authenticator) or 'manual' (it could not; the user finishes with the code).
    auto?: 'working' | 'approve' | 'manual';
    approveNumber?: string;
    autoMessage?: string;
}

export interface SignInOnPageOptions {
    username: string;
    password: string;
    // Page that requires the sign-in; Microsoft returns here when done (e.g. SharePoint).
    startUrl: string;
    timeoutMs?: number;
    onMfa?: (prompt: MfaPrompt | null) => void;
    // Returns an MFA code typed by the user (SMS / app code), once.
    takeMfaCode?: () => string | null;
    // Called when the sign-in moves to a new tab (the previous one was closed).
    onPage?: (page: Page) => void;
    // Called every few seconds and on each new screen to save the session obtained so far,
    // so it survives Edge being closed from outside.
    checkpoint?: () => Promise<unknown>;
    // Device-code sign-in: the code to type at login.microsoft.com/device. The sign-in ends
    // when isDone() says so (the token arrived in Node), not when a target page loads.
    deviceCode?: string;
    isDone?: () => boolean;
    // Answer to "Stay signed in?". Default yes; the device flow answers no (the session lives
    // in Node, and a persistent browser session is what managed Edge reacts to).
    staySignedIn?: boolean;
}

/** Microsoft rejected the e-mail or password; retrying will not help. */
export class MicrosoftCredentialsError extends Error {}

/** The whole browser was closed (not just the tab), e.g. by an Edge policy on the VM. */
export class BrowserClosedError extends Error {
    constructor() {
        super('El navegador de Microsoft se cerró por completo (¿política de Edge en la VM?).');
    }
}

const LOGIN_HOSTS = /(^|\.)login\.(microsoftonline\.com|live\.com|microsoft\.com)$/i;
// Pages Microsoft sends you to after signing in when there is no return address.
const LANDING_HOSTS = /(^|\.)(m365\.cloud\.microsoft|office\.com|microsoft365\.com|office365\.com)$/i;

/** The number to type in Authenticator, read from the page text as a fallback. */
async function numberFromPage(page: Page): Promise<string> {
    const body = (await page.locator('body').innerText({ timeout: 1000 }).catch(() => '')) ?? '';
    return /(?:number|número)[^\d]{0,200}?\b(\d{2,3})\b/i.exec(body)?.[1] ?? '';
}

/** Where a stuck sign-in stopped: URL and visible text (nothing is saved to disk). */
async function describe(page: Page): Promise<string> {
    if (page.isClosed()) {
        return 'La página se cerró.';
    }
    const body = ((await page.locator('body').innerText({ timeout: 2000 }).catch(() => '')) ?? '')
        .replace(/\s+/g, ' ')
        .slice(0, 300);
    return `Se quedó en ${page.url().split('?')[0]} mostrando: "${body}".`;
}

async function visible(page: Page, selector: string): Promise<boolean> {
    // Any visible match: Microsoft pages keep hidden copies of some fields before the real one.
    return page.locator(selector).filter({ visible: true }).first().isVisible().catch(() => false);
}

/** Types into the visible field and reports whether the value stuck (never submit an empty field). */
async function typeInto(page: Page, selector: string, value: string): Promise<boolean> {
    const field = page.locator(`${selector}:visible`).first();
    await field.fill(value, { timeout: 5000 }).catch(() => undefined);
    return (await field.inputValue({ timeout: 2000 }).catch(() => '')) === value;
}

async function text(page: Page, selector: string): Promise<string> {
    const shown = page.locator(selector).filter({ visible: true }).first();
    const any = page.locator(selector).first();
    const value = (await shown.count().catch(() => 0)) ? shown : any;
    return ((await value.textContent({ timeout: 500 }).catch(() => '')) ?? '').trim();
}

/**
 * Signs in to Microsoft 365 inside an existing page, reacting to whatever screen Microsoft
 * shows (account picker, e-mail, password, MFA, "Stay signed in?"). MFA is reported through
 * `onMfa` instead of needing a visible window: the number to match, or a request for a code.
 */
export async function signInOnPage(page: Page, options: SignInOnPageOptions): Promise<Page> {
    const timeoutMs = options.timeoutMs ?? 180000;
    const deadline = Date.now() + timeoutMs;
    const context = page.context();
    let current = page;
    let timer: NodeJS.Timeout | undefined;
    // Hard stop: the loop checks the deadline, but a stuck browser call must not keep the
    // request open forever.
    const hardStop = new Promise<never>((_, reject) => {
        timer = setTimeout(async () => {
            reject(new Error(`El inicio de sesión de Microsoft no terminó a tiempo. ${await describe(current)}`));
        }, timeoutMs + 15000);
    });
    try {
        for (let attempt = 1; attempt <= 4; attempt++) {
            // No single browser action may block the loop for long.
            current.setDefaultTimeout(10000);
            const result = await Promise.race([signInLoop(current, options, deadline), hardStop]);
            if (result === 'done') {
                return current;
            }
            // The tab was closed under us. On managed machines Edge can move work sites to the
            // user's own profile right after MFA. The session cookies are already in this
            // context, so continue in a new tab: SharePoint signs in silently.
            console.log('[MICROSOFT LOGIN] La pestaña se cerró; se continúa en una nueva con la sesión ya iniciada.');
            current = await context.newPage().catch(() => {
                throw new BrowserClosedError();
            });
            options.onPage?.(current);
        }
        throw new Error('Edge cerró la pestaña de inicio de sesión varias veces seguidas.');
    } finally {
        clearTimeout(timer);
        options.onMfa?.(null);
    }
}

async function signInLoop(page: Page, options: SignInOnPageOptions, deadline: number): Promise<'done' | 'closed'> {
    try {
        return await signInSteps(page, options, deadline);
    } catch (error) {
        // Any browser call fails once the tab is gone; report the closing instead.
        if (page.isClosed()) {
            return 'closed';
        }
        throw error;
    }
}

async function signInSteps(page: Page, options: SignInOnPageOptions, deadline: number): Promise<'done' | 'closed'> {
    const target = new URL(options.startUrl).hostname;
    let lastHeartbeat = Date.now();
    let emailSentAt = 0;
    let passwordSentAt = 0;
    let lastPrompt = '';
    let onTargetSince = 0;
    let lastStep = '';
    let stepSince = Date.now();
    let checkpointAt = 0;
    let codeSentAt = 0;
    let confirmedAt = 0;

    const report = (prompt: MfaPrompt | null) => {
        const key = JSON.stringify(prompt);
        if (key !== lastPrompt) {
            lastPrompt = key;
            options.onMfa?.(prompt);
        }
    };

    await page.goto(options.startUrl, { waitUntil: 'commit', timeout: 60000 }).catch(() => undefined);

    try {
        while (Date.now() < deadline) {
            if (options.isDone?.()) {
                return 'done';
            }
            if (page.isClosed()) {
                return 'closed';
            }
            const url = new URL(page.url());
            if (!options.isDone && url.hostname === target) {
                // Signed in once SharePoint has set its session cookie. Without it, only
                // accept the page after it stayed put (it may still redirect to the login).
                const cookies = await page.context().cookies(options.startUrl);
                if (cookies.some((c) => /^(FedAuth|rtFa)$/i.test(c.name))) {
                    return 'done';
                }
                onTargetSince = onTargetSince || Date.now();
                if (Date.now() - onTargetSince > 5000) {
                    return 'done';
                }
            } else {
                onTargetSince = 0;
            }
            const onLogin = LOGIN_HOSTS.test(url.hostname);

            // Log each new screen so a stuck sign-in can be diagnosed from the server console.
            const step = `${url.hostname}${url.pathname}`;
            if (step !== lastStep) {
                lastStep = step;
                stepSince = Date.now();
                console.log(`[MICROSOFT LOGIN] ${step}`);
                checkpointAt = 0;
            }
            if (options.checkpoint && Date.now() - checkpointAt > 3000) {
                checkpointAt = Date.now();
                await options.checkpoint().catch(() => undefined);
            }
            const stuckFor = Date.now() - stepSince;
            if (Date.now() - lastHeartbeat > 15000) {
                // Same screen for a while: log what it shows.
                lastHeartbeat = Date.now();
                const shown = ((await page.locator('body').innerText({ timeout: 2000 }).catch(() => '')) ?? '')
                    .replace(/\s+/g, ' ')
                    .slice(0, 160);
                console.log(`[MICROSOFT LOGIN] esperando en ${step} (${Math.round(stuckFor / 1000)}s): "${shown}"`);
            }

            if (onLogin && (await visible(page, '#passwordError, #usernameError'))) {
                const message = (await text(page, '#passwordError, #usernameError')) || '';
                if (/please enter|enter your|escrib|introduzc|ingres/i.test(message)) {
                    // "Please enter your password": the field was sent empty (e.g. a click on
                    // the live screenshot pressed "Sign in" first). Not wrong credentials: re-send.
                    if (Date.now() - passwordSentAt > 8000 && Date.now() - emailSentAt > 8000) {
                        console.log(`[MICROSOFT LOGIN] Campo vacío ("${message}"); se vuelve a escribir.`);
                        emailSentAt = 0;
                        passwordSentAt = 0;
                    }
                } else {
                    throw new MicrosoftCredentialsError(message || 'Microsoft rechazó el correo o la contraseña.');
                }
            }

            // Device-code screens: type the code, then confirm the app.
            if (options.deviceCode && onLogin) {
                const shown = ((await page.locator('body').innerText({ timeout: 1000 }).catch(() => '')) ?? '').replace(/\s+/g, ' ');
                if (/code (you entered )?(has )?expired|(is not|isn.t) valid|ha expirado|no es v[aá]lido/i.test(shown) && !(await visible(page, 'input[type="password"]'))) {
                    throw new Error(`Microsoft no aceptó el código: "${shown.slice(0, 160)}"`);
                }
                if (Date.now() - codeSentAt > 8000 && (await visible(page, '#otc'))) {
                    console.log('[MICROSOFT LOGIN] Escribiendo el código de dispositivo');
                    if (await typeInto(page, '#otc', options.deviceCode)) {
                        await page.click('#idSIButton9, input[type="submit"]').catch(() => undefined);
                    }
                    codeSentAt = Date.now();
                    await page.waitForTimeout(400).catch(() => undefined);
                    continue;
                }
                if (
                    Date.now() - confirmedAt > 8000 &&
                    /are you trying to sign in|est[aá]s? intentando iniciar sesi[oó]n|intentas iniciar sesi[oó]n/i.test(shown)
                ) {
                    console.log('[MICROSOFT LOGIN] Confirmando el inicio de sesión de la aplicación');
                    await page.click('#idSIButton9, input[type="submit"]').catch(() => undefined);
                    confirmedAt = Date.now();
                    await page.waitForTimeout(400).catch(() => undefined);
                    continue;
                }
            }

            // "Stay signed in?" first: after MFA it is served at /SAS/ProcessAuth, which
            // would otherwise look like the MFA screen. Recognised by its content.
            const kmsi =
                onLogin &&
                (/\/kmsi/i.test(url.pathname) ||
                    (await visible(page, '#KmsiCheckboxField, #KmsiDescription')) ||
                    (await visible(page, 'text=/Stay signed in\\?|mantener la sesi[oó]n iniciada/i')));

            // MFA screens next. Recognised by their content: Microsoft serves them at
            // /SAS/… or at /<tenant>/login depending on the flow.
            if (kmsi) {
                const stay = options.staySignedIn !== false;
                console.log(`[MICROSOFT LOGIN] Mantener sesión iniciada: ${stay ? 'Sí' : 'No'}`);
                await page
                    .locator(
                        stay
                            ? '#idSIButton9, button[data-testid="primaryButton"], input[type="submit"][value="Yes"], input[type="submit"][value="Sí"], button:has-text("Yes"), button:has-text("Sí")'
                            : '#idBtn_Back, button[data-testid="secondaryButton"], input[type="button"][value="No"], input[type="submit"][value="No"], button:has-text("No")'
                    )
                    .first()
                    .click({ timeout: 3000 })
                    .catch(() => undefined);
                // The persistent sign-in cookie arrives with this answer: save it at once.
                await page.waitForLoadState('domcontentloaded', { timeout: 3000 }).catch(() => undefined);
                await options.checkpoint?.().catch(() => undefined);
                checkpointAt = Date.now();
            } else if (onLogin && (await visible(page, '#idTxtBx_SAOTCC_OTC'))) {
                // MFA by code (SMS or authenticator code).
                report({
                    kind: 'code',
                    number: '',
                    message: (await text(page, '#idDiv_SAOTCC_Description')) || 'Escribe el código de verificación.'
                });
                const code = options.takeMfaCode?.();
                if (code) {
                    await page.fill('#idTxtBx_SAOTCC_OTC', code);
                    await page.click('#idSubmit_SAOTCC_Continue');
                }
            } else if (
                onLogin &&
                ((await visible(page, '#idRichContext_DisplaySign, #idDiv_SAOTCAS_Title, #idDiv_SAOTCAS_Description')) ||
                    (await visible(page, 'text=/Approve sign.in request|Aprobar (la )?solicitud de inicio de sesi/i')) ||
                    /\/SAS\/(BeginAuth|EndAuth)/i.test(url.pathname))
            ) {
                // MFA by approval, usually with a number to type in Authenticator.
                report({
                    kind: 'approve',
                    number: (await text(page, '#idRichContext_DisplaySign, .displaySign')) || (await numberFromPage(page)),
                    message:
                        (await text(page, '#idDiv_SAOTCAS_Description')) ||
                        'Aprueba la solicitud de inicio de sesión en Microsoft Authenticator.'
                });
            } else if (onLogin && (await visible(page, '#tilesHolder'))) {
                // "Pick an account": choose ours, or "Use another account".
                const tile = page.locator('#tilesHolder [role="listitem"], #tilesHolder .table').filter({ hasText: options.username });
                if (await tile.count()) {
                    await tile.first().click().catch(() => undefined);
                } else {
                    await page.locator('#otherTile, #otherTileText').first().click().catch(() => undefined);
                }
            } else if (onLogin && Date.now() - emailSentAt > 8000 && (await visible(page, 'input[type="email"]'))) {
                console.log('[MICROSOFT LOGIN] Enviando correo');
                if (await typeInto(page, 'input[type="email"]', options.username)) {
                    await page.click('#idSIButton9, input[type="submit"]');
                }
                emailSentAt = Date.now();
            } else if (onLogin && Date.now() - passwordSentAt > 8000 && (await visible(page, 'input[type="password"]'))) {
                console.log('[MICROSOFT LOGIN] Enviando contraseña');
                if (await typeInto(page, 'input[type="password"]', options.password)) {
                    await page.click('#idSIButton9, input[type="submit"]');
                }
                passwordSentAt = Date.now();
            } else if (
                !options.isDone &&
                !onLogin &&
                url.hostname !== target &&
                (LANDING_HOSTS.test(url.hostname) || url.protocol === 'chrome-error:' || stuckFor > 20000)
            ) {
                // Signed in but landed on Microsoft 365 home (or stuck elsewhere): go to the
                // target. Other hosts (e.g. Windows seamless sign-on) are waited for.
                await page.goto(options.startUrl, { waitUntil: 'commit', timeout: 60000 }).catch(() => undefined);
                stepSince = Date.now();
            }
            await page.waitForTimeout(400).catch(() => undefined);
        }
        throw new Error(`El inicio de sesión de Microsoft no terminó a tiempo. ${await describe(page)}`);
    } finally {
        report(null);
    }
}
