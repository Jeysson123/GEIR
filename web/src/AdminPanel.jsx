import React from 'react';
import { api } from './api';
import Icon from './icons';

var DRAFT_KEY = 'geir.adminDraft';
var REQUIREMENT = /\b\d{4}-[A-Za-z]+-\d+(?:-\d+)?\b/g;

function loadDraft() {
    try {
        return JSON.parse(window.sessionStorage.getItem(DRAFT_KEY)) || {};
    } catch (e) {
        return {};
    }
}

function saveDraft(draft) {
    try {
        window.sessionStorage.setItem(DRAFT_KEY, JSON.stringify(draft));
    } catch (e) {
        // Only a convenience: the form still works without it.
    }
}

function size(bytes) {
    return bytes >= 1024 * 1024 ? (bytes / 1024 / 1024).toFixed(1) + ' MB' : Math.max(1, Math.round(bytes / 1024)) + ' KB';
}

function shortDate(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    return m ? m[3] + '/' + m[2] + '/' + m[1] : '';
}

/** "Usuario Uno (Empresa) <Usuario_Uno@example.com>" → "Usuario Uno". */
function recipientName(address) {
    var name = address.replace(/<[^>]*>/, '').replace(/\([^)]*\)/, '').replace(/"/g, '').trim();
    return name || address.replace(/[<>]/g, '');
}

var BUILD_BADGE = { compilando: 'badge-edit', listo: 'badge-new', error: 'badge-warn', omitido: 'badge-warn' };
var BUILD_REASON = {
    cambio: 'cambiado en la rama',
    depende: 'usa un proyecto cambiado',
    falta: 'librería que falta en Maven'
};

var KINDS = [
    { kind: 'geir', title: 'Documento GEIR', icon: 'sheet' },
    { kind: 'derivado', title: 'Documentos para aprobación', icon: 'folderOpen' },
    { kind: 'sprint', title: 'Lista de historias del sprint', icon: 'table' },
    { kind: 'correo', title: 'Correos que no llegaron a Outlook (.eml)', icon: 'mailDraft' }
];

/**
 * Third card: relates requirements to the user stories, fills the GEIR (request IDs,
 * work plan), derives the approval workbooks and the sprint list, and prepares e-mail
 * drafts next to the GEIR. Nothing is sent.
 */
export default class AdminPanel extends React.Component {
    constructor(props) {
        super(props);
        var draft = loadDraft();
        this.state = {
            cambio: draft.cambio || '',
            geir: draft.geir || '',
            sprint: draft.sprint || '',
            fechaPase: draft.fechaPase || '',
            requerimientos: draft.requerimientos || '',
            rama: draft.rama || '',
            carpeta: draft.carpeta || '',
            folder: { folder: '', editable: false },
            busy: false,
            build: null,
            buildError: '',
            startingBuild: false,
            error: '',
            result: null,
            downloading: ''
        };
        this.generate = this.generate.bind(this);
        this.compile = this.compile.bind(this);
        this.poll = null;
    }

    componentDidMount() {
        api.adminBuildDefaults().then(
            (folder) => this.setState({ folder: folder }),
            () => undefined
        );
    }

    componentWillUnmount() {
        clearTimeout(this.poll);
    }

    change(name, value) {
        var patch = {};
        patch[name] = value;
        this.setState(patch, () => {
            saveDraft({
                cambio: this.state.cambio,
                geir: this.state.geir,
                sprint: this.state.sprint,
                fechaPase: this.state.fechaPase,
                requerimientos: this.state.requerimientos,
                rama: this.state.rama,
                carpeta: this.state.carpeta
            });
        });
    }

    generate(e) {
        e.preventDefault();
        if (this.state.busy || !this.props.location) {
            return;
        }
        this.props.onWork(true);
        this.setState({ busy: true, error: '', result: null });
        api.adminGenerate({
            location: this.props.location,
            cambio: this.state.cambio,
            geir: this.state.geir,
            sprint: this.state.sprint,
            fechaPase: this.state.fechaPase,
            requerimientos: this.state.requerimientos
        }).then(
            (result) => {
                this.props.onWork(false);
                this.setState({ busy: false, result: result });
                this.props.onDocument(result.document);
                if (this.state.rama.trim()) {
                    this.startBuild();
                }
            },
            (error) => {
                this.props.onWork(false);
                this.setState({ busy: false, error: error.message });
            }
        );
    }

    compile(e) {
        e.preventDefault();
        this.startBuild();
    }

    // Compilation runs on the server for minutes: start it, then follow it every 2 seconds.
    startBuild() {
        if (this.state.startingBuild || this.buildRunning() || !this.props.location || !this.state.rama.trim()) {
            return;
        }
        clearTimeout(this.poll);
        this.setState({ startingBuild: true, buildError: '', build: null });
        api.adminBuild({
            location: this.props.location,
            rama: this.state.rama.trim(),
            carpeta: this.state.carpeta.trim(),
            requerimientos: this.state.requerimientos
        }).then(
            (job) => {
                this.setState({ startingBuild: false, build: job });
                this.follow(job.id);
            },
            (error) => this.setState({ startingBuild: false, buildError: error.message })
        );
    }

    follow(id) {
        this.poll = setTimeout(() => {
            api.adminBuildStatus(id).then(
                (job) => {
                    this.setState({ build: job });
                    if (job.state === 'en cola' || job.state === 'trabajando') {
                        this.follow(id);
                    }
                },
                (error) => this.setState({ buildError: error.message })
            );
        }, 2000);
    }

    buildRunning() {
        var b = this.state.build;
        return !!b && (b.state === 'en cola' || b.state === 'trabajando');
    }

    download(name) {
        this.setState({ downloading: name });
        api.adminDownload(name).then(
            () => this.setState({ downloading: '' }),
            (error) => this.setState({ downloading: '', error: error.message })
        );
    }

    field(name, label, icon, placeholder, type) {
        return (
            <div className="field">
                <label htmlFor={'adm-' + name}>
                    <Icon name={icon} /> {label}
                </label>
                <input
                    id={'adm-' + name}
                    type={type || 'text'}
                    className={type ? '' : 'mono'}
                    placeholder={placeholder}
                    value={this.state[name]}
                    onChange={(e) => this.change(name, e.target.value)}
                />
            </div>
        );
    }

    renderResult() {
        var r = this.state.result;
        if (!r) {
            return null;
        }
        var plan = r.plan;
        var drafts = r.drafts || [];
        var inbox = drafts.filter((d) => d.saved).length;
        return (
            <div className="admin-result">
                <div className="alert alert-ok alert-icon">
                    <Icon name="checkCircle" size={18} />
                    <span>
                        Listo: {r.files.length} archivos en la carpeta del GEIR ({r.names.cambio} · {r.names.geir} · Sprint{' '}
                        {r.names.sprint})
                        {inbox ? (
                            <span>
                                {' '}y {inbox} borradores en tu Outlook (carpeta <b>Borradores</b>). Revísalos y envíalos tú.
                            </span>
                        ) : (
                            '.'
                        )}
                    </span>
                </div>

                {r.warnings.length ? (
                    <div className="alert alert-warn alert-icon">
                        <Icon name="alert" size={18} />
                        <ul className="warnings">
                            {r.warnings.map((w, i) => (
                                <li key={i}>{w}</li>
                            ))}
                        </ul>
                    </div>
                ) : null}

                <h3 className="section-title">
                    <Icon name="link" /> Requerimientos ↔ historias de usuario
                </h3>
                <div className="table-wrap">
                    <table className="data-table">
                        <thead>
                            <tr>
                                <th>Requerimiento</th>
                                <th>Historia</th>
                                <th>Descripción</th>
                                <th>Tipo</th>
                                <th>Tareas del GEIR</th>
                            </tr>
                        </thead>
                        <tbody>
                            {r.matches.map((m) => (
                                <tr key={m.requirement}>
                                    <td className="mono nowrap">{m.requirement}</td>
                                    <td className="mono">{m.storyId}</td>
                                    <td>{m.title}</td>
                                    <td className="nowrap">
                                        <span className={'badge ' + (/correcci/i.test(m.type) ? 'badge-warn' : 'badge-edit')}>{m.type}</span>
                                    </td>
                                    <td>
                                        {m.tasks.length ? (
                                            m.tasks.map((t) => (
                                                <span key={t.id} className="mini-chip" title={t.how === 'hija' ? 'Tarea hija de la historia' : 'Es la misma historia'}>
                                                    <Icon name={t.how === 'hija' ? 'git' : 'check'} size={12} />
                                                    {t.number ? '#' + t.number + ' ' : ''}
                                                    {t.id}
                                                </span>
                                            ))
                                        ) : (
                                            <span className="muted small">sin tareas</span>
                                        )}
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>

                <h3 className="section-title">
                    <Icon name="calendar" /> Plan de trabajo ({plan.totalDays} días hábiles)
                </h3>
                <p className="muted small">
                    Desde la primera historia{plan.first ? ' (' + plan.first.id + ', ' + shortDate(plan.first.date) + ')' : ''} hasta la
                    última agregada{plan.last ? ' (' + plan.last.id + ', ' + shortDate(plan.last.date) + ')' : ''}. Pase a producción:{' '}
                    <b>{shortDate(plan.release)}</b>.
                </p>
                <div className="table-wrap">
                    <table className="data-table">
                        <thead>
                            <tr>
                                <th>Actividad</th>
                                <th className="num">Días</th>
                                <th>Inicio</th>
                                <th>Fin</th>
                            </tr>
                        </thead>
                        <tbody>
                            {plan.rows.map((p) => (
                                <tr key={p.activity} className={p.days ? '' : 'row-muted'}>
                                    <td>{p.activity}</td>
                                    <td className="num">{p.days || '—'}</td>
                                    <td>{shortDate(p.start)}</td>
                                    <td>{shortDate(p.end)}</td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>

                {drafts.length ? (
                    <div>
                        <h3 className="section-title">
                            <Icon name="mailDraft" /> Borradores en tu Outlook (no se envían)
                        </h3>
                        <div className="file-group">
                            {drafts.map((d) => (
                                <div key={d.subject} className="file-row">
                                    <Icon name={d.saved ? 'mail' : 'alert'} className="file-icon" />
                                    <span className="file-name" title={d.to.join('; ')}>
                                        {d.subject}
                                        {d.to.length ? <span className="muted"> · para {d.to.map(recipientName).join(', ')}</span> : null}
                                    </span>
                                    {d.saved ? (
                                        d.webLink ? (
                                            <a className="btn btn-secondary btn-sm" href={d.webLink} target="_blank" rel="noopener noreferrer">
                                                <Icon name="external" /> Abrir en Outlook
                                            </a>
                                        ) : (
                                            <span className="badge badge-edit">En Borradores</span>
                                        )
                                    ) : (
                                        <button
                                            type="button"
                                            className="btn btn-secondary btn-sm"
                                            disabled={this.state.downloading === d.file}
                                            onClick={() => this.download(d.file)}
                                        >
                                            <Icon name={this.state.downloading === d.file ? 'loader' : 'download'} spin={this.state.downloading === d.file} />
                                            .eml
                                        </button>
                                    )}
                                </div>
                            ))}
                        </div>
                    </div>
                ) : null}

                <h3 className="section-title">
                    <Icon name="folder" /> Archivos generados
                </h3>
                <div className="file-groups">
                    {KINDS.map((k) => {
                        var files = r.files.filter((f) => f.kind === k.kind);
                        if (!files.length) {
                            return null;
                        }
                        return (
                            <div key={k.kind} className="file-group">
                                <div className="file-group-title">
                                    <Icon name={k.icon} /> {k.title}
                                </div>
                                {files.map((f) => (
                                    <div key={f.name} className="file-row">
                                        <Icon name={k.kind === 'correo' ? 'mail' : 'sheet'} className="file-icon" />
                                        <span className="file-name">{f.name}</span>
                                        <span className="muted small">{size(f.size)}</span>
                                        <button
                                            type="button"
                                            className="btn btn-secondary btn-sm"
                                            disabled={this.state.downloading === f.name}
                                            onClick={() => this.download(f.name)}
                                        >
                                            <Icon name={this.state.downloading === f.name ? 'loader' : 'download'} spin={this.state.downloading === f.name} />
                                            Descargar
                                        </button>
                                    </div>
                                ))}
                            </div>
                        );
                    })}
                </div>
            </div>
        );
    }

    renderBuild() {
        var b = this.state.build;
        if (!b) {
            return null;
        }
        var running = b.state === 'en cola' || b.state === 'trabajando';
        return (
            <div className="admin-result build-result">
                <h3 className="section-title">
                    <Icon name="git" /> Compilación de <span className="mono">{b.branch}</span>
                    {b.base ? <span className="muted small"> comparada con {b.base}</span> : null}
                </h3>
                {running ? (
                    <div className="progress">
                        <Icon name="loader" spin /> {b.step}
                    </div>
                ) : b.state === 'listo' ? (
                    <div className="alert alert-ok alert-icon">
                        <Icon name="checkCircle" size={18} />
                        <span>
                            Listo: {b.files.length} archivos en <b>{b.deploy}</b>
                        </span>
                    </div>
                ) : (
                    <div className="alert alert-error alert-icon">
                        <Icon name="alert" size={18} />
                        <span>{b.error}</span>
                    </div>
                )}

                {b.warnings.length ? (
                    <div className="alert alert-warn alert-icon">
                        <Icon name="alert" size={18} />
                        <ul className="warnings">
                            {b.warnings.map((w, i) => (
                                <li key={i}>{w}</li>
                            ))}
                        </ul>
                    </div>
                ) : null}

                {b.ids && b.ids.length ? (
                    <p className="muted small">
                        Historias de la GEIR buscadas en los commits:{' '}
                        {b.ids.map((id) => (
                            <span key={id} className="mini-chip">
                                {id}
                            </span>
                        ))}
                    </p>
                ) : null}

                {b.commits.length ? (
                    <div>
                        <h3 className="section-title">
                            <Icon name="list" /> Commits de la rama: {b.commits.filter((c) => c.related.length).length} de la GEIR
                            <span className="muted small"> (de {b.commits.length}; los demás no se compilan)</span>
                        </h3>
                        <div className="table-wrap build-commits">
                            <table className="data-table">
                                <thead>
                                    <tr>
                                        <th>Commit</th>
                                        <th>Historia</th>
                                        <th>Mensaje</th>
                                        <th>Proyectos</th>
                                        <th>Autor</th>
                                        <th>Fecha</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {b.commits.map((c) => (
                                        <tr key={c.sha} className={c.related.length ? '' : 'row-muted'}>
                                            <td className="mono nowrap">{c.short}</td>
                                            <td className="mono nowrap">{c.related.length ? c.related.join(', ') : '—'}</td>
                                            <td>
                                                {c.merge ? <span className="badge">merge</span> : null} {c.title}
                                            </td>
                                            <td className="mono small">{c.projects && c.projects.length ? c.projects.join(', ') : '—'}</td>
                                            <td className="nowrap">{c.author}</td>
                                            <td className="nowrap">{shortDate(c.date)}</td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>
                ) : null}

                {b.projects.length ? (
                    <div>
                        <h3 className="section-title">
                            <Icon name="code" /> Proyectos compilados
                        </h3>
                        <div className="table-wrap">
                            <table className="data-table">
                                <thead>
                                    <tr>
                                        <th>Proyecto</th>
                                        <th>Por qué</th>
                                        <th>Comando</th>
                                        <th>Estado</th>
                                        <th>Deploy</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {b.projects.map((p) => (
                                        <tr key={p.dir + p.name}>
                                            <td>
                                                <b>{p.name}</b>
                                                <div className="muted small mono">{p.dir || '(raíz)'}</div>
                                            </td>
                                            <td className="small">{BUILD_REASON[p.reason]}</td>
                                            <td className="mono small">{p.command}</td>
                                            <td className="nowrap">
                                                <span className={'badge ' + (BUILD_BADGE[p.state] || '')}>
                                                    {p.state === 'compilando' ? <Icon name="loader" spin size={12} /> : null} {p.state}
                                                </span>
                                            </td>
                                            <td className="small mono">
                                                {p.outputs.length ? (
                                                    p.outputs.map((o) => <div key={o}>{o}</div>)
                                                ) : p.destination ? (
                                                    p.destination
                                                ) : (
                                                    <span className="muted">solo se instala</span>
                                                )}
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>
                ) : null}

                {b.files.length ? (
                    <div>
                        <h3 className="section-title">
                            <Icon name="folder" /> Carpeta Deploy
                        </h3>
                        <div className="file-group">
                            {b.files.map((f) => (
                                <div key={f.path} className="file-row">
                                    <Icon name={f.error ? 'alert' : 'code'} className="file-icon" />
                                    <span className="file-name mono" title={f.error || f.path}>
                                        {f.path}
                                        {f.error ? <span className="muted"> · no se subió: {f.error}</span> : null}
                                    </span>
                                    {f.error ? <span className="badge badge-warn">falló</span> : <span className="muted small">{size(f.size)}</span>}
                                </div>
                            ))}
                        </div>
                    </div>
                ) : null}

                <details className="build-log" open={b.state === 'error'}>
                    <summary>Registro ({b.log.length} líneas)</summary>
                    <pre>{b.log.join('\n')}</pre>
                </details>
            </div>
        );
    }

    render() {
        var count = (this.state.requerimientos.match(REQUIREMENT) || []).length;
        var enabled = !!this.props.location;
        return (
            <section className={'panel admin-panel' + (enabled ? '' : ' panel-disabled')}>
                <div className="panel-head">
                    <span className="step step-icon">
                        <Icon name="shield" size={16} />
                    </span>
                    <div>
                        <h2>Opciones de administrador</h2>
                        <p className="muted">
                            Relaciona requerimientos con historias, completa el GEIR y prepara los documentos de aprobación y los
                            borradores de correo en tu Outlook. Ningún correo se envía.
                        </p>
                    </div>
                </div>
                <form onSubmit={this.generate}>
                    <div className="admin-grid">
                        {this.field('cambio', 'Nombre del cambio', 'hash', 'ej. CRQ12345')}
                        {this.field('geir', 'Nombre de la GEIR', 'sheet', 'ej. GEIR_XX000-20260101-000000')}
                        {this.field('sprint', 'Sprint', 'calendar', 'ej. 24-2025')}
                        {this.field('fechaPase', 'Fecha de pase (opcional)', 'clock', '', 'date')}
                    </div>
                    <div className="build-row">
                        <div className="field">
                            <label htmlFor="adm-rama">
                                <Icon name="git" /> Rama (compilar binarios para Deploy)
                            </label>
                            <input
                                id="adm-rama"
                                type="text"
                                className="mono"
                                placeholder="ej. release/CRQ12345-sprint-1"
                                value={this.state.rama}
                                onChange={(e) => this.change('rama', e.target.value)}
                            />
                        </div>
                        <button
                            type="button"
                            className="btn btn-secondary"
                            onClick={this.compile}
                            disabled={!enabled || !this.state.rama.trim() || this.state.startingBuild || this.buildRunning()}
                        >
                            <Icon
                                name={this.state.startingBuild || this.buildRunning() ? 'loader' : 'zap'}
                                spin={this.state.startingBuild || this.buildRunning()}
                            />
                            {this.buildRunning() ? 'Compilando…' : 'Compilar binarios'}
                        </button>
                    </div>
                    {this.state.folder.editable ? (
                        <div className="field">
                            <label htmlFor="adm-carpeta">
                                <Icon name="folderOpen" /> Carpeta local del repositorio
                            </label>
                            <input
                                id="adm-carpeta"
                                type="text"
                                className="mono"
                                placeholder={this.state.folder.folder || 'ej. C:\\dev\\mi-repositorio'}
                                value={this.state.carpeta}
                                onChange={(e) => this.change('carpeta', e.target.value)}
                            />
                        </div>
                    ) : null}
                    <p className="hint">
                        <Icon name="git" size={13} />{' '}
                        {this.state.carpeta.trim() ? (
                            <span>
                                Se compila en <b className="mono">{this.state.carpeta.trim()}</b> (fetch + pull de la rama).
                            </span>
                        ) : this.state.folder.folder ? (
                            <span>
                                Se compila en <b className="mono">{this.state.folder.folder}</b> (fetch + pull de la rama).
                            </span>
                        ) : (
                            <span>No se encontró una carpeta local del repositorio: se descargará de GitLab.</span>
                        )}
                    </p>
                    <p className="hint">
                        <Icon name="info" size={13} /> Lo que dejes vacío se crea como <b>TEMPORAL</b>. Sin fecha de pase se usa la de
                        "Liberación y Pase a Producción" del plan. Con una rama se hace pull, se compilan los proyectos que tocan sus
                        commits (Maven y <span className="mono">npm run buildprod</span>) y los binarios quedan en la carpeta{' '}
                        <b>Deploy</b> junto al GEIR.
                    </p>
                    <div className="field">
                        <label htmlFor="adm-requerimientos">
                            <Icon name="list" /> Requerimientos e historias de usuario
                            {count ? <span className="badge badge-edit count-badge">{count} detectados</span> : null}
                        </label>
                        <textarea
                            id="adm-requerimientos"
                            className="mono"
                            rows={6}
                            placeholder={
                                'Un requerimiento y su historia por línea, por ejemplo:\n2026-REQ-0101 100002\n2026-REQ-0102: 100001 - Botones separados en el formulario\n(o pega las filas del Excel del sprint, o un arreglo JSON)'
                            }
                            value={this.state.requerimientos}
                            onChange={(e) => this.change('requerimientos', e.target.value)}
                        />
                    </div>
                    <div className="row admin-actions">
                        <button className="btn btn-primary" type="submit" disabled={!enabled || this.state.busy || !count}>
                            <Icon name={this.state.busy ? 'loader' : 'sparkles'} spin={this.state.busy} />
                            {this.state.busy ? 'Generando…' : 'Generar documentos y borradores'}
                        </button>
                        {!enabled ? <span className="muted small">Primero abre el documento GEIR (paso 1).</span> : null}
                    </div>
                </form>
                {this.state.busy ? (
                    <div className="progress">
                        <Icon name="clock" /> Relacionando historias, calculando el plan y creando los archivos…{' '}
                        <span className="elapsed">{this.props.seconds}s</span>
                    </div>
                ) : null}
                {this.state.error ? (
                    <div className="alert alert-error alert-icon">
                        <Icon name="alert" size={18} />
                        <span>{this.state.error}</span>
                    </div>
                ) : null}
                {this.renderResult()}
                {this.state.buildError ? (
                    <div className="alert alert-error alert-icon">
                        <Icon name="alert" size={18} />
                        <span>{this.state.buildError}</span>
                    </div>
                ) : null}
                {this.renderBuild()}
            </section>
        );
    }
}
