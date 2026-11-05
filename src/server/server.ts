import express, { NextFunction, Request, Response } from 'express';
import * as path from 'path';
import { Environment } from '../infrastructure/config/environment';
import {
    UserError,
    deleteTask,
    login,
    lookupTask,
    microsoftStatus,
    openDocument,
    saveTask
} from '../application/geir.service';
import { Session, closeBrowser, destroySession, getSession } from '../application/session';
import { VERSION } from './version';
import { adminDefaults, adminFile, adminGenerate } from '../application/admin.service';
import { buildDefaults, buildStatus, startBuild } from '../application/build.service';
import { configAsset } from '../domain/admin/config';

const WEB_ROOT = path.resolve(__dirname, '../../web');

type AuthedRequest = Request & { session?: Session };

const app = express();
app.disable('x-powered-by');
if (Environment.geir.trustProxy) {
    app.set('trust proxy', true);
}
app.use(express.json({ limit: '60mb' }));
app.use((_req, res, next) => {
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('X-Frame-Options', 'DENY');
    res.setHeader('Referrer-Policy', 'no-referrer');
    next();
});

// Logo from config/assets: logo.local.png (git-ignored, an installation's own brand) or logo.png.
app.get('/brand/logo.png', (_req, res) => {
    const logo = configAsset('assets/logo.local.png') ?? configAsset('assets/logo.png');
    if (!logo) {
        res.status(404).end();
        return;
    }
    res.set('Cache-Control', 'no-cache').type('image/png').send(logo);
});
// no-cache: the browser always checks for a newer interface (e.g. after copying new files).
app.use(express.static(WEB_ROOT, { index: 'index.html', setHeaders: (res) => res.setHeader('Cache-Control', 'no-cache') }));

function tokenOf(req: Request): string | undefined {
    const header = req.header('authorization') ?? '';
    return header.startsWith('Bearer ') ? header.slice(7) : undefined;
}

function requireSession(req: AuthedRequest, res: Response, next: NextFunction) {
    const session = getSession(tokenOf(req));
    if (!session) {
        res.status(401).json({ error: 'Tu sesión expiró. Vuelve a iniciar sesión.' });
        return;
    }
    req.session = session;
    // A new Microsoft refresh token goes to this user's browser only (sessionStorage); the
    // server never writes it to disk. Empty header = forget the one stored.
    const send = res.json.bind(res);
    res.json = (body?: any) => {
        if (session.microsoftTokenChanged && !res.headersSent) {
            session.microsoftTokenChanged = false;
            res.setHeader('X-GEIR-Microsoft', session.microsoftRefreshToken ?? '');
        }
        return send(body);
    };
    next();
}

// Login attempts per address: 10 every 15 minutes (it checks network passwords).
const attempts = new Map<string, { count: number; since: number }>();
function limitLogin(req: Request, res: Response, next: NextFunction) {
    const key = req.ip ?? 'unknown';
    const now = Date.now();
    const entry = attempts.get(key);
    if (!entry || now - entry.since > 15 * 60 * 1000) {
        attempts.set(key, { count: 1, since: now });
    } else if (++entry.count > 10) {
        res.status(429).json({ error: 'Demasiados intentos de inicio de sesión. Espera 15 minutos.' });
        return;
    }
    next();
}

const handle =
    (fn: (req: AuthedRequest, res: Response) => Promise<unknown>) =>
    (req: AuthedRequest, res: Response, next: NextFunction) =>
        fn(req, res).catch(next);

app.post(
    '/api/auth/login',
    limitLogin,
    handle(async (req, res) => {
        const email = String(req.body?.email ?? '').trim();
        // Microsoft token kept by this tab from an earlier login, only for the same mailbox.
        const kept = req.body?.microsoft;
        const refresh =
            kept && typeof kept.refreshToken === 'string' && String(kept.email ?? '').toLowerCase() === email.toLowerCase()
                ? kept.refreshToken
                : undefined;
        const session = await login({ username: req.body?.username, password: req.body?.password, email, microsoftRefreshToken: refresh });
        attempts.delete(req.ip ?? 'unknown');
        res.json({
            token: session.token,
            user: { username: session.credentials.username, displayName: session.displayName, email: session.names.email },
            mock: Environment.geir.mock,
            version: VERSION
        });
    })
);

app.get('/api/auth/me', requireSession, (req: AuthedRequest, res) => {
    res.json({
        user: { username: req.session!.credentials.username, displayName: req.session!.displayName, email: req.session!.names.email },
        mock: Environment.geir.mock,
        version: VERSION
    });
});

app.post(
    '/api/auth/logout',
    handle(async (req, res) => {
        const token = tokenOf(req);
        if (token) {
            await destroySession(token);
        }
        res.json({ ok: true });
    })
);

// Polled by the UI while a document operation runs, to show Microsoft MFA requests.
app.get('/api/microsoft/status', requireSession, (req: AuthedRequest, res) => {
    res.json(microsoftStatus(req.session!));
});




app.post(
    '/api/document/open',
    requireSession,
    handle(async (req, res) => {
        res.json(await openDocument(req.session!, String(req.body?.location ?? '')));
    })
);

app.post(
    '/api/tasks/lookup',
    requireSession,
    handle(async (req, res) => {
        const { location, taskId, refresh } = req.body ?? {};
        res.json(await lookupTask(req.session!, String(location ?? ''), String(taskId ?? ''), Boolean(refresh)));
    })
);

app.post(
    '/api/tasks/save',
    requireSession,
    handle(async (req, res) => {
        const { location, task, evidenceChanged } = req.body ?? {};
        res.json(await saveTask(req.session!, String(location ?? ''), task, Boolean(evidenceChanged)));
    })
);

app.post(
    '/api/tasks/delete',
    requireSession,
    handle(async (req, res) => {
        const { location, taskId } = req.body ?? {};
        res.json(await deleteTask(req.session!, String(location ?? ''), String(taskId ?? '')));
    })
);

// ---------------------------------------------------------------- admin options

app.get(
    '/api/admin/defaults',
    requireSession,
    handle(async (_req, res) => {
        res.json(adminDefaults());
    })
);

app.post(
    '/api/admin/generate',
    requireSession,
    handle(async (req, res) => {
        const b = req.body ?? {};
        res.json(
            await adminGenerate(req.session!, {
                location: String(b.location ?? ''),
                cambio: String(b.cambio ?? ''),
                geir: String(b.geir ?? ''),
                sprint: String(b.sprint ?? ''),
                requerimientos: String(b.requerimientos ?? ''),
                fechaPase: String(b.fechaPase ?? '')
            })
        );
    })
);

app.post(
    '/api/admin/build',
    requireSession,
    handle(async (req, res) => {
        const b = req.body ?? {};
        res.json(
            startBuild(req.session!, {
                location: String(b.location ?? ''),
                rama: String(b.rama ?? ''),
                carpeta: String(b.carpeta ?? ''),
                requerimientos: String(b.requerimientos ?? '')
            })
        );
    })
);

app.get('/api/admin/build-defaults', requireSession, (_req, res) => {
    res.json(buildDefaults());
});

app.get('/api/admin/build/:id', requireSession, (req: AuthedRequest, res) => {
    const job = buildStatus(req.session!, String(req.params.id ?? ''));
    if (!job) {
        res.status(404).json({ error: 'Esa compilación ya no existe.' });
        return;
    }
    res.json(job);
});

app.get('/api/admin/files/:name', requireSession, (req: AuthedRequest, res) => {
    const name = String(req.params.name ?? '');
    const data = adminFile(req.session!, name);
    if (!data) {
        res.status(404).json({ error: 'Ese archivo ya no está disponible; vuelve a generar.' });
        return;
    }
    const type = /\.eml$/i.test(name)
        ? 'message/rfc822'
        : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    res.set('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(name)}`).type(type).send(data);
});

app.use((error: any, _req: Request, res: Response, _next: NextFunction) => {
    const status = error instanceof UserError ? error.status : 500;
    if (status >= 500) {
        console.error('[GEIR]', error);
    }
    res.status(status).json({ error: error?.message ?? String(error) });
});

const server = app.listen(Environment.geir.port, Environment.geir.host, () => {
    const shown = ['0.0.0.0', '::'].includes(Environment.geir.host) ? 'localhost' : Environment.geir.host;
    console.log(`GEIR ${VERSION} listo en http://${shown}:${Environment.geir.port}${Environment.geir.mock ? '  (modo demostración)' : ''}`);
    console.log(
        `Escuchando en ${Environment.geir.host}:${Environment.geir.port} · rutas locales ${Environment.geir.allowLocalPaths ? 'permitidas' : 'bloqueadas'} · navegador ${Environment.geir.browserChannel}`
    );
    console.log(`Carpeta: ${path.resolve(__dirname, '../..')}`);
});
server.on('error', (error: any) => {
    if (error?.code === 'EADDRINUSE') {
        console.error(
            `\n[GEIR] El puerto ${Environment.geir.port} ya está en uso: hay OTRO GEIR (u otro programa) corriendo.\n` +
                '       Ciérralo (Ctrl+C en su consola) o ejecuta en PowerShell:\n' +
                `       Get-NetTCPConnection -LocalPort ${Environment.geir.port} -State Listen | % { Stop-Process -Id $_.OwningProcess -Force }\n` +
                '       y vuelve a ejecutar npm start.\n'
        );
        process.exit(1);
    }
    throw error;
});

async function shutdown() {
    server.close();
    await closeBrowser();
    process.exit(0);
}
process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);
