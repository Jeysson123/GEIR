import { APIRequestContext, request } from '@playwright/test';

/**
 * Microsoft 365 access without a browser: OAuth device-code sign-in. GEIR shows a code,
 * types it (with the user's e-mail and password) at microsoft.com/devicelogin and gets a
 * token for SharePoint, renewed on its own and also used for Microsoft Graph (Outlook).
 *
 * Several people use the same server, so nothing is written to disk: the refresh token
 * lives in this object (one per GEIR session) and in the sessionStorage of that user's
 * browser tab, which sends it back after a new login (see `onRefreshToken`).
 */

// Microsoft Office: public client, pre-authorized for SharePoint. Override with MICROSOFT_CLIENT_ID.
const OFFICE_CLIENT_ID = 'd3590ed6-52b3-4102-aeff-aad2292ab01c';

export interface DevicePrompt {
    userCode: string;
    verificationUri: string;
    message: string;
}

interface TokenResponse {
    access_token?: string;
    refresh_token?: string;
    expires_in?: number;
    device_code?: string;
    user_code?: string;
    verification_uri?: string;
    message?: string;
    interval?: number;
    error?: string;
    error_description?: string;
}

/**
 * A Node-side HTTP client (no browser) of its own: cookies set by Microsoft never mix
 * between users. Honors HTTPS_PROXY / HTTP_PROXY.
 */
export function newHttpClient(): Promise<APIRequestContext> {
    const proxy = process.env.HTTPS_PROXY || process.env.https_proxy || process.env.HTTP_PROXY || process.env.http_proxy;
    return request.newContext({ proxy: proxy ? { server: proxy } : undefined, timeout: 60000 });
}

export class MicrosoftToken {
    // Access tokens by resource: the same refresh token also gives Microsoft Graph (Outlook drafts).
    private access = new Map<string, { token: string; expiresAt: number }>();
    private refreshToken: string | null;
    private client: Promise<APIRequestContext> | null = null;

    /**
     * @param resource e.g. https://contoso-my.sharepoint.com
     * @param refreshToken one kept by the user's browser from an earlier sign-in, if any.
     * @param onRefreshToken called with every new refresh token (null when it stops working).
     */
    constructor(
        private resource: string,
        refreshToken: string | null,
        private onRefreshToken: (token: string | null) => void = () => undefined
    ) {
        this.refreshToken = refreshToken || null;
    }

    /** This session's HTTP client, for SharePoint and Graph too. */
    http(): Promise<APIRequestContext> {
        if (!this.client) {
            this.client = newHttpClient();
        }
        return this.client;
    }

    async dispose() {
        const client = this.client;
        this.client = null;
        this.access.clear();
        this.refreshToken = null;
        await (await client?.catch(() => null))?.dispose();
    }

    private get clientId() {
        return process.env.MICROSOFT_CLIENT_ID || OFFICE_CLIENT_ID;
    }

    private scope(resource = this.resource) {
        return `${resource}/.default offline_access`;
    }

    /**
     * A valid access token, renewed from the refresh token; null when a sign-in is needed.
     * @param resource another Microsoft 365 API, e.g. https://graph.microsoft.com (default: SharePoint).
     */
    async token(resource = this.resource): Promise<string | null> {
        const cached = this.access.get(resource);
        if (cached && Date.now() < cached.expiresAt - 60000) {
            return cached.token;
        }
        if (!this.refreshToken) {
            return null;
        }
        const res = await this.post('token', {
            grant_type: 'refresh_token',
            refresh_token: this.refreshToken,
            client_id: this.clientId,
            scope: this.scope(resource)
        });
        if (!res.access_token) {
            const reason = res.error_description?.split(/\r?\n/)[0] || res.error || 'sin respuesta';
            if (resource !== this.resource && res.error && res.error !== 'invalid_grant') {
                // The sign-in is fine but this API is not allowed (e.g. consent): do not lose the session.
                throw new Error(`Microsoft no dio acceso a ${resource}: ${reason}`);
            }
            console.log(`[MICROSOFT TOKEN] No se pudo renovar (${reason}); hay que iniciar sesión.`);
            this.forget();
            return null;
        }
        this.accept(res, resource);
        return this.access.get(resource)!.token;
    }

    /** Device-code sign-in: reports the code to show, then waits until it is entered. */
    async signIn(onPrompt: (prompt: DevicePrompt) => void, timeoutMs = 600000): Promise<void> {
        const code = await this.post('devicecode', { client_id: this.clientId, scope: this.scope() });
        if (!code.device_code || !code.user_code) {
            throw new Error(`Microsoft no entregó un código de inicio de sesión: ${code.error_description || code.error}`);
        }
        onPrompt({
            userCode: code.user_code,
            verificationUri: code.verification_uri || 'https://microsoft.com/devicelogin',
            message: code.message || ''
        });
        let interval = (code.interval || 5) * 1000;
        const deadline = Date.now() + Math.min(timeoutMs, (code.expires_in || 900) * 1000);
        while (Date.now() < deadline) {
            await new Promise((resolve) => setTimeout(resolve, interval));
            const res = await this.post('token', {
                grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
                client_id: this.clientId,
                device_code: code.device_code
            });
            if (res.access_token) {
                this.accept(res);
                console.log('[MICROSOFT TOKEN] Sesión iniciada.');
                return;
            }
            if (res.error === 'authorization_pending') {
                continue;
            }
            if (res.error === 'slow_down') {
                interval += 5000;
                continue;
            }
            throw new Error(
                res.error === 'authorization_declined'
                    ? 'Se rechazó el inicio de sesión en Microsoft.'
                    : `Microsoft no aceptó el inicio de sesión: ${res.error_description || res.error}`
            );
        }
        throw new Error('El código de Microsoft venció sin que se escribiera. Intenta de nuevo.');
    }

    forget() {
        this.access.clear();
        if (this.refreshToken) {
            this.refreshToken = null;
            this.onRefreshToken(null);
        }
    }

    private accept(res: TokenResponse, resource = this.resource) {
        this.access.set(resource, { token: res.access_token!, expiresAt: Date.now() + (res.expires_in || 3600) * 1000 });
        if (res.refresh_token && res.refresh_token !== this.refreshToken) {
            this.refreshToken = res.refresh_token;
            this.onRefreshToken(res.refresh_token);
        }
    }

    private async post(endpoint: 'token' | 'devicecode', form: Record<string, string>): Promise<TokenResponse> {
        const tenant = process.env.MICROSOFT_TENANT || 'organizations';
        const res = await (await this.http()).post(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/${endpoint}`, { form });
        return (await res.json().catch(() => ({ error: `HTTP ${res.status()}` }))) as TokenResponse;
    }
}
