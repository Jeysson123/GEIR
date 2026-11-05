// Everything GEIR keeps in the browser lives in sessionStorage: it is per tab, survives
// reloads and is deleted when the tab closes. Several people use the same server and maybe
// the same computer, so nothing is left in localStorage. Browser storage can throw (private
// mode, blocked site data), so every access is guarded.

var SESSION_KEY = 'geir.session';
var DOCUMENT_KEY = 'geir.documentPath';
var ACCOUNT_KEY = 'geir.account';
var MICROSOFT_KEY = 'geir.microsoft';

function read(key) {
    try {
        return window.sessionStorage.getItem(key);
    } catch (e) {
        return null;
    }
}

function write(key, value) {
    try {
        if (value === null) {
            window.sessionStorage.removeItem(key);
        } else {
            window.sessionStorage.setItem(key, value);
        }
    } catch (e) {
        // Not fatal: the app keeps working for this tab.
    }
}

function readJson(key) {
    try {
        return JSON.parse(read(key)) || null;
    } catch (e) {
        return null;
    }
}

// The login session: only the server token, never the password.
export function loadSession() {
    return readJson(SESSION_KEY);
}

export function saveSession(session) {
    write(SESSION_KEY, session ? JSON.stringify(session) : null);
}

export function loadDocumentPath() {
    return read(DOCUMENT_KEY) || '';
}

export function saveDocumentPath(path) {
    write(DOCUMENT_KEY, path || null);
}

// User name and Microsoft e-mail prefill the login of this tab (never the password).
export function loadAccount() {
    return readJson(ACCOUNT_KEY) || {};
}

export function saveAccount(username, email) {
    write(ACCOUNT_KEY, JSON.stringify({ username: username, email: email }));
}

// Microsoft 365 sign-in of this user ({ email, refreshToken }). The server keeps it only in
// memory; this copy lets a new login in the same tab (e.g. after a server restart) skip the
// Microsoft code. Deleted on logout and when the tab closes.
export function loadMicrosoft() {
    return readJson(MICROSOFT_KEY);
}

export function saveMicrosoft(value) {
    write(MICROSOFT_KEY, value ? JSON.stringify(value) : null);
}
