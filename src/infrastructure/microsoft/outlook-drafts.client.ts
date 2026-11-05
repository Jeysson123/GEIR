import { APIRequestContext } from '@playwright/test';

/**
 * Outlook drafts through Microsoft Graph: the message is saved in the Drafts folder of the
 * signed-in user's own mailbox. Graph only creates it; nothing here sends e-mail (there is
 * no call to /send anywhere in GEIR).
 */

export const GRAPH = 'https://graph.microsoft.com';

// Graph rejects requests over 4 MB; the MIME text grows by a third when encoded in base64.
const MAX_MIME_BYTES = 2.9 * 1024 * 1024;

export interface SavedDraft {
    id: string;
    subject: string;
    webLink: string;
}

/** The token was not accepted (expired or revoked): sign in again and retry. */
export class GraphAuthError extends Error {}

export class OutlookDraftsClient {
    constructor(private bearer: string, private http: APIRequestContext) {}

    /** Creates a draft from a complete MIME message (headers, HTML, inline logo, attachments). */
    async saveMime(mime: Buffer): Promise<SavedDraft> {
        if (mime.length > MAX_MIME_BYTES) {
            throw new Error(
                `el correo pesa ${(mime.length / 1024 / 1024).toFixed(1)} MB y Outlook acepta hasta ${(MAX_MIME_BYTES / 1024 / 1024).toFixed(1)} MB por esta vía`
            );
        }
        const res = await this.http.post(`${GRAPH}/v1.0/me/messages`, {
            headers: { Authorization: `Bearer ${this.bearer}`, 'Content-Type': 'text/plain' },
            data: mime.toString('base64')
        });
        const body: any = await res.json().catch(() => ({}));
        if (res.status() === 401) {
            throw new GraphAuthError(body?.error?.message || 'token rechazado');
        }
        if (!res.ok()) {
            const reason = body?.error?.message || `HTTP ${res.status()}`;
            throw new Error(res.status() === 403 ? `sin permiso para crear borradores en el buzón (${reason})` : reason);
        }
        return { id: body.id, subject: body.subject ?? '', webLink: body.webLink ?? '' };
    }
}
