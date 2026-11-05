import React from 'react';
import { api } from './api';
import { loadDocumentPath, saveDocumentPath } from './storage';
import AdminPanel from './AdminPanel';
import Icon, { BrandLogo } from './icons';
import TaskEditor from './TaskEditor';
import { VERSION } from './version';

function Elapsed(props) {
    return <span className="elapsed">{props.seconds}s</span>;
}

export default class Dashboard extends React.Component {
    constructor(props) {
        super(props);
        this.state = {
            path: loadDocumentPath(),
            openedPath: '',
            document: null,
            docBusy: false,
            docError: '',
            taskId: '',
            lookup: null, // result of /api/tasks/lookup
            evidenceChanged: false,
            lookupBusy: false,
            lookupError: '',
            saveBusy: false,
            deleting: false,
            docMessage: '',
            saveMessage: '',
            saveError: '',
            seconds: 0,
            msSigningIn: false,
            mfa: null, // { kind: 'approve' | 'code', number, message } while Microsoft asks for MFA
            codeCopied: false
        };
        this.openDocument = this.openDocument.bind(this);
        this.lookup = this.lookup.bind(this);
        this.save = this.save.bind(this);
        this.remove = this.remove.bind(this);
        this.changeTask = this.changeTask.bind(this);
    }

    componentDidMount() {
        // "Si ya existe en el storage buscarla": reopen the remembered document.
        if (this.state.path) {
            this.openDocument();
        }
    }

    componentWillUnmount() {
        clearInterval(this.timer);
        clearInterval(this.mfaTimer);
    }

    startTimer() {
        clearInterval(this.timer);
        clearInterval(this.mfaTimer);
        this.setState({ seconds: 0 });
        this.timer = setInterval(() => this.setState({ seconds: this.state.seconds + 1 }), 1000);
        // Microsoft may ask for MFA while the document is read or saved.
        this.mfaTimer = setInterval(() => {
            api.microsoftStatus().then(
                (status) => {
                    this.setState({ msSigningIn: status.signingIn, mfa: status.mfa });
                },
                () => null
            );
        }, 2000);
    }

    stopTimer() {
        clearInterval(this.timer);
        clearInterval(this.mfaTimer);
        this.setState({ msSigningIn: false, mfa: null, codeCopied: false });
    }

    renderMicrosoftSignIn() {
        const { msSigningIn, mfa } = this.state;
        if (!msSigningIn && !mfa) {
            return null;
        }
        return (
            <section className="panel mfa-panel" role="status">
                <h2 className="title-icon">
                    <Icon name="shield" size={18} /> Inicio de sesión en Microsoft 365
                </h2>
                {mfa && mfa.kind === 'device' ? (
                    <div>
                        {mfa.auto === 'approve' ? (
                            <div>
                                <p className="title-icon">
                                    <Icon name="phone" size={18} /> Abre <b>Microsoft Authenticator</b> en tu teléfono y aprueba el inicio de sesión
                                    {mfa.approveNumber ? ' escribiendo este número:' : '.'}
                                </p>
                                {mfa.approveNumber ? (
                                    <div className="mfa-number" aria-label={'Número ' + mfa.approveNumber}>
                                        {mfa.approveNumber}
                                    </div>
                                ) : null}
                                <p className="hint">GEIR ya escribió el código, tu correo y tu contraseña. Solo falta tu aprobación.</p>
                            </div>
                        ) : mfa.auto === 'working' ? (
                            <p className="muted title-icon">
                                <Icon name="loader" spin /> {mfa.autoMessage}
                            </p>
                        ) : mfa.autoMessage ? (
                            <div className="alert alert-warn">{mfa.autoMessage}</div>
                        ) : null}
                        {mfa.auto === 'manual' || !mfa.auto ? (
                            this.renderManualDevice(mfa)
                        ) : (
                            <details className="ms-manual">
                                <summary>¿No avanza? Hazlo tú con el código</summary>
                                {this.renderManualDevice(mfa)}
                            </details>
                        )}
                    </div>
                ) : (
                    <p className="muted title-icon">
                        <Icon name="loader" spin /> Conectando con Microsoft 365…
                    </p>
                )}
            </section>
        );
    }

    renderManualDevice(mfa) {
        return (
            <div className="ms-manual-body">
                <p>Pulsa el botón: se copia el código y se abre Microsoft en tu navegador.</p>
                <div className="mfa-number mfa-code" aria-label={'Código ' + mfa.number}>
                    {mfa.number}
                </div>
                <button className="btn btn-primary" type="button" onClick={() => this.openDeviceLogin(mfa)}>
                    <Icon name={this.state.codeCopied ? 'check' : 'copy'} />
                    {this.state.codeCopied ? 'Código copiado: pégalo en Microsoft (Ctrl+V)' : 'Copiar código y abrir Microsoft'}
                </button>
                <p className="hint">Pega el código, elige tu cuenta y acepta. GEIR continúa solo.</p>
            </div>
        );
    }

    openDeviceLogin(mfa) {
        this.copyCode(mfa.number);
        window.open(mfa.url, '_blank', 'noopener');
    }

    copyCode(code) {
        const done = () => this.setState({ codeCopied: true });
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(code).then(done, () => this.copyWithSelection(code, done));
        } else {
            this.copyWithSelection(code, done);
        }
    }

    copyWithSelection(code, done) {
        const area = document.createElement('textarea');
        area.value = code;
        document.body.appendChild(area);
        area.select();
        try {
            if (document.execCommand('copy')) {
                done();
            }
        } catch (e) {
            // The code stays visible to type by hand.
        }
        document.body.removeChild(area);
    }

    openDocument(e) {
        if (e) {
            e.preventDefault();
        }
        const path = this.state.path.trim();
        if (!path || this.state.docBusy) {
            return;
        }
        this.startTimer();
        this.setState({ docBusy: true, docError: '', saveMessage: '', docMessage: '' });
        api.openDocument(path).then(
            (doc) => {
                this.stopTimer();
                saveDocumentPath(path);
                const switched = path !== this.state.openedPath;
                this.setState({
                    docBusy: false,
                    document: doc,
                    openedPath: path,
                    lookup: switched ? null : this.state.lookup
                });
            },
            (error) => {
                this.stopTimer();
                this.setState({ docBusy: false, docError: error.message });
            }
        );
    }

    lookup(e, refresh, idOverride) {
        if (e) {
            e.preventDefault();
        }
        const taskId = String(idOverride || this.state.taskId).trim();
        if (!taskId || !this.state.document || this.state.lookupBusy) {
            return;
        }
        this.startTimer();
        this.setState({ taskId: taskId, lookupBusy: true, lookupError: '', saveMessage: '', saveError: '', docMessage: '' });
        if (!refresh) {
            this.setState({ lookup: null });
        }
        api.lookupTask(this.state.openedPath, taskId, refresh).then(
            (result) => {
                this.stopTimer();
                this.setState({
                    lookupBusy: false,
                    lookup: result,
                    // New tasks write the drafted evidence. Edits keep the file's picture unless a
                    // regenerate brought images from the Azure discussion.
                    evidenceChanged: result.mode === 'new' ? !!result.task.evidence : !!refresh && result.images.length > 0
                });
            },
            (error) => {
                this.stopTimer();
                this.setState({ lookupBusy: false, lookupError: error.message });
            }
        );
    }

    changeTask(patch, evidenceChanged) {
        const lookup = Object.assign({}, this.state.lookup, { task: Object.assign({}, this.state.lookup.task, patch) });
        this.setState({ lookup: lookup, evidenceChanged: evidenceChanged || this.state.evidenceChanged, saveMessage: '' });
    }

    save() {
        const lookup = this.state.lookup;
        if (!lookup || this.state.saveBusy) {
            return;
        }
        this.startTimer();
        this.setState({ saveBusy: true, saveError: '', saveMessage: '' });
        api.saveTask(this.state.openedPath, lookup.task, this.state.evidenceChanged).then(
            (result) => {
                this.stopTimer();
                const task = Object.assign({}, lookup.task, { number: result.number });
                this.setState({
                    saveBusy: false,
                    document: result.document,
                    evidenceChanged: false,
                    lookup: Object.assign({}, lookup, { mode: 'edit', task: task }),
                    saveMessage:
                        result.result === 'added'
                            ? 'Tarea ' + task.id + ' agregada como #' + result.number + ' en las cuatro hojas.'
                            : 'Tarea ' + task.id + ' actualizada en el documento.'
                });
            },
            (error) => {
                this.stopTimer();
                this.setState({ saveBusy: false, saveError: error.message });
            }
        );
    }

    remove(target) {
        const lookup = this.state.lookup;
        const task = target || (lookup && lookup.mode === 'edit' ? lookup.task : null);
        if (!task || this.state.saveBusy) {
            return;
        }
        const ok = window.confirm(
            '¿Eliminar la tarea ' +
                task.id +
                (task.title ? ' (' + task.title + ')' : '') +
                ' del documento?\n\nSe quita de las cuatro hojas y las tareas siguientes se renumeran.'
        );
        if (!ok) {
            return;
        }
        this.startTimer();
        this.setState({ saveBusy: true, deleting: true, saveError: '', saveMessage: '', docMessage: '', docError: '' });
        api.deleteTask(this.state.openedPath, task.id).then(
            (result) => {
                this.stopTimer();
                const open = this.state.lookup;
                const closeEditor = open && open.task.id === task.id;
                this.setState({
                    saveBusy: false,
                    deleting: false,
                    document: result.document,
                    lookup: closeEditor ? null : open,
                    taskId: closeEditor ? '' : this.state.taskId,
                    evidenceChanged: false,
                    docMessage: 'Tarea ' + task.id + ' eliminada del documento.'
                });
            },
            (error) => {
                this.stopTimer();
                this.setState({ saveBusy: false, deleting: false, saveError: error.message, docError: error.message });
            }
        );
    }

    renderDocumentPanel() {
        const { path, document, docBusy, docError, openedPath, seconds } = this.state;
        const dirty = document && path.trim() !== openedPath;
        return (
            <section className="panel">
                <div className="panel-head">
                    <span className="step step-icon">
                        <Icon name="sheet" size={16} />
                    </span>
                    <div>
                        <h2>Documento GEIR en OneDrive</h2>
                        <p className="muted">
                            Enlace de la carpeta o del archivo de Excel. Si la carpeta no tiene “N - GEIR.xlsx”, se crea.
                        </p>
                    </div>
                </div>
                <form className="row" onSubmit={this.openDocument}>
                    <input
                        className="grow mono"
                        placeholder="https://empresa-my.sharepoint.com/…  o  C:\ruta\N - GEIR.xlsx"
                        value={path}
                        onChange={(e) => this.setState({ path: e.target.value })}
                    />
                    <button className="btn btn-primary" type="submit" disabled={docBusy || !path.trim()}>
                        <Icon name={docBusy ? 'loader' : document && !dirty ? 'refresh' : 'folderOpen'} spin={docBusy} />
                        {docBusy ? 'Abriendo…' : dirty ? 'Usar esta ruta' : document ? 'Recargar' : 'Abrir'}
                    </button>
                </form>
                {docBusy ? (
                    <div className="progress">
                        <Icon name="cloud" /> Abriendo el documento en OneDrive… <Elapsed seconds={seconds} />
                        {seconds > 5 && !this.state.msSigningIn ? (
                            <div className="hint">La primera vez puede tardar mientras se inicia sesión en Microsoft 365.</div>
                        ) : null}
                    </div>
                ) : null}
                {docError ? (
                    <div className="alert alert-error alert-icon">
                        <Icon name="alert" size={18} />
                        <span>{docError}</span>
                    </div>
                ) : null}
                {this.state.deleting ? (
                    <div className="progress">
                        <Icon name="trash" /> Eliminando la tarea del documento… <Elapsed seconds={seconds} />
                    </div>
                ) : null}
                {this.state.docMessage ? (
                    <div className="alert alert-ok alert-icon">
                        <Icon name="checkCircle" size={18} />
                        <span>{this.state.docMessage}</span>
                    </div>
                ) : null}
                {document ? (
                    <div className="doc-summary">
                        <div className="doc-line">
                            <Icon name="sheet" className="doc-icon" size={18} />
                            <strong>{document.name}</strong>
                            {document.created ? <span className="badge badge-new">Creado ahora</span> : null}
                            {document.webUrl ? (
                                <a className="btn btn-link" href={document.webUrl} target="_blank" rel="noopener noreferrer">
                                    <Icon name="external" /> Abrir en Excel Online
                                </a>
                            ) : null}
                        </div>
                        <div className="muted small mono title-icon">
                            <Icon name="folder" size={13} /> {document.location}
                        </div>
                        {document.tasks.length ? (
                            <div className="task-chips">
                                <span className="muted small title-icon">
                                    <Icon name="list" size={13} /> Tareas en el documento:
                                </span>
                                {document.tasks.map((t) => (
                                    <span className="chip-group" key={t.id + '-' + t.number}>
                                        <button
                                            className="chip"
                                            title={'Editar ' + t.id + ': ' + t.title}
                                            onClick={() => this.lookup(null, false, t.id)}
                                        >
                                            {t.number ? <b>#{t.number}</b> : null} {t.id}
                                            {t.requirement ? <span className="chip-req">{t.requirement}</span> : null}
                                            <span className="chip-title">{t.title}</span>
                                        </button>
                                        <button
                                            className="chip-delete"
                                            title={'Eliminar ' + t.id + ' del documento'}
                                            aria-label={'Eliminar ' + t.id + ' del documento'}
                                            disabled={this.state.saveBusy}
                                            onClick={() => this.remove(t)}
                                        >
                                            <Icon name="x" size={13} />
                                        </button>
                                    </span>
                                ))}
                            </div>
                        ) : (
                            <div className="muted small">El documento todavía no tiene tareas.</div>
                        )}
                    </div>
                ) : null}
            </section>
        );
    }

    renderTaskPanel() {
        const { document, taskId, lookupBusy, lookupError, lookup, seconds } = this.state;
        return (
            <section className={'panel' + (document ? '' : ' panel-disabled')}>
                <div className="panel-head">
                    <span className="step step-icon">
                        <Icon name="task" size={16} />
                    </span>
                    <div>
                        <h2>Tarea de Azure DevOps</h2>
                        <p className="muted">
                            Si el ID no está en el documento se arma con Azure y GitLab; si ya está, se abre para editarlo.
                        </p>
                    </div>
                </div>
                <form className="row" onSubmit={(e) => this.lookup(e, false)}>
                    <input
                        className="id-input mono"
                        inputMode="numeric"
                        placeholder="ID, ej. 100007"
                        disabled={!document}
                        value={taskId}
                        onChange={(e) => this.setState({ taskId: e.target.value.replace(/[^\d]/g, '') })}
                    />
                    <button className="btn btn-primary" type="submit" disabled={!document || lookupBusy || !taskId}>
                        <Icon name={lookupBusy ? 'loader' : 'search'} spin={lookupBusy} />
                        {lookupBusy ? 'Buscando…' : 'Buscar'}
                    </button>
                    {lookup && lookup.mode === 'edit' ? (
                        <button
                            type="button"
                            className="btn btn-secondary"
                            disabled={lookupBusy}
                            onClick={() => this.lookup(null, true)}
                            title="Vuelve a generar los textos desde Azure DevOps y GitLab"
                        >
                            <Icon name="refresh" /> Regenerar desde Azure/GitLab
                        </button>
                    ) : null}
                </form>
                {lookupBusy ? (
                    <div className="progress">
                        <Icon name="git" /> Consultando el documento, Azure DevOps y GitLab… <Elapsed seconds={seconds} />
                        <div className="hint">Lee la tarea, su discusión (imágenes) y los archivos modificados.</div>
                    </div>
                ) : null}
                {lookupError ? (
                    <div className="alert alert-error alert-icon">
                        <Icon name="alert" size={18} />
                        <span>{lookupError}</span>
                    </div>
                ) : null}
            </section>
        );
    }

    render() {
        const { session, onLogout } = this.props;
        const { lookup, saveBusy, saveMessage, saveError, document, seconds } = this.state;
        return (
            <div className="app">
                <header className="topbar">
                    <div className="brand small-brand">
                        <BrandLogo size={34} />
                        <strong className="app-title">
                            GEIR
                        </strong>
                        {session.mock ? <span className="badge badge-warn">Modo demostración</span> : null}
                        <span className="muted small" title="Versión de la interfaz">v{VERSION}</span>
                        {session.version !== VERSION ? (
                            <span className="badge badge-warn" title="Reinicia el servidor y pulsa Ctrl+F5">
                                {'Servidor ' + (session.version || 'antiguo') + ' ≠ interfaz: reinicia npm start y pulsa Ctrl+F5'}
                            </span>
                        ) : null}
                    </div>
                    <div className="topbar-right">
                        <span className="user-chip" title={session.user.username}>
                            <span className="avatar">
                                <Icon name="user" size={15} />
                            </span>
                            {session.user.displayName || session.user.username}
                        </span>
                        <button className="btn btn-ghost btn-logout" onClick={onLogout} title="Cerrar sesión">
                            <Icon name="logout" /> <span className="logout-text">Cerrar sesión</span>
                        </button>
                    </div>
                </header>
                <main className="content">
                    {this.renderMicrosoftSignIn()}
                    {this.renderDocumentPanel()}
                    {this.renderTaskPanel()}
                    <AdminPanel
                        location={document ? this.state.openedPath : ''}
                        seconds={seconds}
                        onWork={(on) => (on ? this.startTimer() : this.stopTimer())}
                        onDocument={(doc) => this.setState({ document: doc })}
                    />
                    {lookup ? (
                        <TaskEditor
                            key={lookup.task.id}
                            lookup={lookup}
                            nextNumber={
                                document ? Math.max(0, ...document.tasks.map((t) => t.number || 0)) + 1 : null
                            }
                            evidenceChanged={this.state.evidenceChanged}
                            onChange={this.changeTask}
                        />
                    ) : null}
                </main>
                {lookup ? (
                    <footer className="savebar">
                        <div className="savebar-status">
                            {saveBusy ? (
                                <span>
                                    {this.state.deleting ? 'Eliminando de ' : 'Guardando en '}
                                    {document ? document.name : 'el documento'}… <Elapsed seconds={seconds} />
                                </span>
                            ) : saveError ? (
                                <span className="text-error">{saveError}</span>
                            ) : saveMessage ? (
                                <span className="text-ok">
                                    {saveMessage}{' '}
                                    {document && document.webUrl ? (
                                        <a href={document.webUrl} target="_blank" rel="noopener noreferrer">
                                            <Icon name="external" size={13} /> Abrir en Excel Online
                                        </a>
                                    ) : null}
                                </span>
                            ) : (
                                <span className="muted">
                                    {lookup.mode === 'new'
                                        ? 'Se agregará en el siguiente espacio disponible de cada hoja.'
                                        : 'Se actualizarán las filas existentes de esta tarea.'}
                                </span>
                            )}
                        </div>
                        <div className="savebar-actions">
                            {lookup.mode === 'edit' ? (
                                <button className="btn btn-danger" disabled={saveBusy} onClick={() => this.remove()}>
                                    <Icon name={this.state.deleting ? 'loader' : 'trash'} spin={this.state.deleting} />
                                    {this.state.deleting ? 'Eliminando…' : 'Eliminar del documento'}
                                </button>
                            ) : null}
                            <button className="btn btn-primary" disabled={saveBusy} onClick={this.save}>
                                <Icon
                                    name={saveBusy && !this.state.deleting ? 'loader' : lookup.mode === 'new' ? 'plus' : 'save'}
                                    spin={saveBusy && !this.state.deleting}
                                />
                                {saveBusy && !this.state.deleting
                                    ? 'Guardando…'
                                    : lookup.mode === 'new'
                                    ? 'Agregar al documento'
                                    : 'Guardar cambios'}
                            </button>
                        </div>
                    </footer>
                ) : null}
            </div>
        );
    }
}
