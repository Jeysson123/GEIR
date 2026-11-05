import React from 'react';
import Icon from './icons';

const TABS = [
    { key: 'description', label: 'Descripción Requerimiento', icon: 'edit' },
    { key: 'design', label: 'Diseño Solución', icon: 'code' },
    { key: 'log', label: 'Bitácora Pruebas Técnicas', icon: 'image' },
    { key: 'uat', label: 'Plan de Pruebas UAT', icon: 'checkCircle' }
];

const RESULTS = ['Satiscatorio', 'Fallido', 'Parcial'];

function imageSrc(image) {
    if (!image) {
        return '';
    }
    const ext = image.extension === 'jpg' ? 'jpeg' : image.extension;
    return 'data:image/' + ext + ';base64,' + image.base64;
}

function Field(props) {
    const id = 'f-' + props.name;
    const value = props.value || '';
    return (
        <div className={'field' + (props.wide ? ' field-wide' : '')}>
            <label htmlFor={id}>
                {props.label}
                {props.required && !value.trim() ? <span className="req">requerido</span> : null}
            </label>
            {props.rows ? (
                <textarea
                    id={id}
                    rows={props.rows}
                    value={value}
                    placeholder={props.placeholder}
                    onChange={(e) => props.onChange(props.name, e.target.value)}
                />
            ) : (
                <input
                    id={id}
                    value={value}
                    placeholder={props.placeholder}
                    className={props.mono ? 'mono' : ''}
                    onChange={(e) => props.onChange(props.name, e.target.value)}
                />
            )}
            {props.help ? <div className="hint">{props.help}</div> : null}
        </div>
    );
}

export default class TaskEditor extends React.Component {
    constructor(props) {
        super(props);
        this.state = { tab: 'description', dragging: false };
        this.set = this.set.bind(this);
        this.onPaste = this.onPaste.bind(this);
        this.onFile = this.onFile.bind(this);
    }

    set(name, value) {
        const patch = {};
        patch[name] = value;
        this.props.onChange(patch);
    }

    setEvidence(image) {
        this.props.onChange({ evidence: image ? { base64: image.base64, extension: image.extension } : null }, true);
    }

    readImageFile(file) {
        if (!file || !/^image\/(png|jpe?g|gif)$/.test(file.type)) {
            return;
        }
        const reader = new FileReader();
        reader.onload = () => {
            const dataUrl = String(reader.result);
            const comma = dataUrl.indexOf(',');
            this.setEvidence({
                base64: dataUrl.slice(comma + 1),
                extension: file.type.split('/')[1].replace('jpg', 'jpeg')
            });
        };
        reader.readAsDataURL(file);
    }

    onPaste(e) {
        const items = (e.clipboardData && e.clipboardData.items) || [];
        for (let i = 0; i < items.length; i++) {
            if (items[i].kind === 'file') {
                e.preventDefault();
                this.readImageFile(items[i].getAsFile());
                return;
            }
        }
    }

    onFile(e) {
        this.readImageFile(e.target.files && e.target.files[0]);
        e.target.value = '';
    }

    renderHeader() {
        const { lookup } = this.props;
        const task = lookup.task;
        const wi = lookup.workItem;
        return (
            <div className="editor-head">
                <div>
                    <div className="editor-title">
                        <span className={'badge ' + (lookup.mode === 'new' ? 'badge-new' : 'badge-edit')}>
                            <Icon name={lookup.mode === 'new' ? 'plus' : 'edit'} size={12} /> {lookup.mode === 'new' ? 'Nueva' : 'Editando'}
                        </span>
                        <span className="mono">{task.id}</span>
                        <span className="sep">·</span>
                        <span>{task.title || 'Sin título'}</span>
                        {task.number ? <span className="muted">(#{task.number} en el documento)</span> : null}
                    </div>
                    {wi ? (
                        <div className="muted small">
                            {[wi.workItemType, wi.state, wi.iterationPath].filter(Boolean).join(' · ')}{' '}
                            <a href={wi.editUrl} target="_blank" rel="noopener noreferrer">
                                <Icon name="external" size={12} /> Ver en Azure
                            </a>
                        </div>
                    ) : null}
                </div>
            </div>
        );
    }

    renderDescription(task) {
        const number = task.number || this.props.nextNumber || 'N';
        return (
            <div className="tab-body">
                <div className="grid">
                    <Field name="title" label="Título de la tarea" value={task.title} onChange={this.set} required wide />
                    <Field
                        name="benefit"
                        label="Beneficio (Para Que)"
                        rows={4}
                        value={task.benefit}
                        onChange={this.set}
                        wide
                        help="Tomado de la descripción en Azure (la parte “para …” de la historia, o su primera oración)."
                    />
                </div>
                <div className="preview">
                    <div className="preview-label">
                        <Icon name="eye" size={13} /> Así quedará en la hoja
                    </div>
                    <div className="preview-row">
                        <span className="preview-key">Yo Quiero:</span>
                        <span>
                            {task.requirement
                                ? number + ' - ' + task.requirement + ': ' + task.id + ' - ' + task.title
                                : number + ' - : ' + task.id + ' : ' + task.title}
                        </span>
                    </div>
                    <div className="preview-row">
                        <span className="preview-key">Para Que:</span>
                        <span>{number} - {task.benefit}</span>
                    </div>
                </div>
            </div>
        );
    }

    renderDesign(task) {
        const { files, sources } = this.props.lookup;
        return (
            <div className="tab-body split">
                <div className="split-main">
                    <Field name="functionalDesign" label="Diseño Funcional" rows={5} value={task.functionalDesign} onChange={this.set} wide />
                    <Field
                        name="technicalDesign"
                        label="Diseño Técnico"
                        rows={9}
                        value={task.technicalDesign}
                        onChange={this.set}
                        wide
                        placeholder="Describe los archivos modificados (backend / frontend)…"
                        help="Generado desde los cambios en GitLab. Revisa y ajusta la redacción antes de guardar."
                    />
                    <Field name="securityDesign" label="Diseño Seguridad" value={task.securityDesign} onChange={this.set} wide />
                </div>
                <aside className="split-side">
                    <h3 className="title-icon">
                        <Icon name="git" /> Cambios en GitLab
                    </h3>
                    {sources.length ? (
                        <ul className="sources">
                            {sources.map((s) => (
                                <li key={s.project + s.ref}>
                                    <a href={s.url} target="_blank" rel="noopener noreferrer">
                                        {s.ref}
                                    </a>{' '}
                                    <span className="muted">{s.title}</span>
                                </li>
                            ))}
                        </ul>
                    ) : null}
                    {files.length ? (
                        <ul className="files">
                            {files.map((f) => (
                                <li key={f.path}>
                                    <div className="file-line">
                                        <span className={'layer layer-' + f.layer.replace(/\s/g, '-')}>{f.layer}</span>
                                        <span className="mono file-name" title={f.path}>
                                            {f.name}
                                        </span>
                                        <span className="stat">
                                            <span className="plus">+{f.additions}</span> <span className="minus">-{f.deletions}</span>
                                        </span>
                                    </div>
                                    <div className="small">{f.sentence}</div>
                                </li>
                            ))}
                        </ul>
                    ) : (
                        <p className="muted small">
                            {this.props.lookup.mode === 'edit' && !this.props.lookup.workItem
                                ? 'Pulsa “Regenerar desde Azure/GitLab” para ver los archivos modificados.'
                                : 'No se encontraron archivos modificados para esta tarea.'}
                        </p>
                    )}
                </aside>
            </div>
        );
    }

    renderLog(task) {
        const { lookup, evidenceChanged } = this.props;
        const current = task.evidence;
        return (
            <div className="tab-body">
                <Field
                    name="workItemUrl"
                    label="Enlace de la tarea (debajo de la evidencia)"
                    value={task.workItemUrl}
                    onChange={this.set}
                    mono
                    wide
                />
                <div className="field field-wide">
                    <label>Evidencia (captura)</label>
                    <div
                        className={'evidence' + (this.state.dragging ? ' evidence-drag' : '')}
                        tabIndex={0}
                        onPaste={this.onPaste}
                        onDragOver={(e) => {
                            e.preventDefault();
                            this.setState({ dragging: true });
                        }}
                        onDragLeave={() => this.setState({ dragging: false })}
                        onDrop={(e) => {
                            e.preventDefault();
                            this.setState({ dragging: false });
                            this.readImageFile(e.dataTransfer.files && e.dataTransfer.files[0]);
                        }}
                    >
                        {current ? (
                            <img src={imageSrc(current)} alt="Evidencia seleccionada" />
                        ) : (
                            <div className="muted">Sin imagen. Pega una captura (Ctrl+V), arrástrala aquí o elige un archivo.</div>
                        )}
                    </div>
                    <div className="row evidence-actions">
                        <label className="btn btn-secondary file-btn">
                            <Icon name="image" /> Elegir archivo…
                            <input type="file" accept="image/png,image/jpeg,image/gif" onChange={this.onFile} />
                        </label>
                        {current ? (
                            <button type="button" className="btn btn-ghost" onClick={() => this.setEvidence(null)}>
                                <Icon name="x" /> Quitar
                            </button>
                        ) : null}
                        <span className="hint">
                            {lookup.mode === 'edit' && !evidenceChanged
                                ? 'Se conserva la imagen que ya tiene el documento.'
                                : current
                                ? 'Se insertará en la Bitácora, ajustada al recuadro.'
                                : ''}
                        </span>
                    </div>
                </div>
                {lookup.images.length ? (
                    <div className="field field-wide">
                        <label>Imágenes de la discusión en Azure ({lookup.images.length})</label>
                        <div className="gallery">
                            {lookup.images.map((img, i) => {
                                const selected = current && current.base64 === img.base64;
                                return (
                                    <button
                                        key={i}
                                        type="button"
                                        className={'thumb' + (selected ? ' thumb-selected' : '')}
                                        onClick={() => this.setEvidence(img)}
                                        title={img.author + ' · ' + new Date(img.date).toLocaleString()}
                                    >
                                        <img src={imageSrc(img)} alt={'Imagen ' + (i + 1)} />
                                        <span className="thumb-meta">
                                            {img.evidenceHint ? <b>Evidencia · </b> : null}
                                            {img.author}
                                        </span>
                                    </button>
                                );
                            })}
                        </div>
                    </div>
                ) : lookup.workItem ? (
                    <p className="muted small">La discusión de la tarea no tiene imágenes.</p>
                ) : null}
            </div>
        );
    }

    renderUat(task) {
        return (
            <div className="tab-body">
                <div className="grid">
                    <Field name="testCase" label="Caso de Prueba" rows={2} value={task.testCase} onChange={this.set} wide />
                    <Field name="given" label="Dado" rows={3} value={task.given} onChange={this.set} />
                    <Field name="when" label="Cuándo" rows={3} value={task.when} onChange={this.set} />
                    <Field name="then" label="Entonces" rows={3} value={task.then} onChange={this.set} wide />
                    <div className="field">
                        <label htmlFor="f-result">Resultado</label>
                        <select id="f-result" value={task.result} onChange={(e) => this.set('result', e.target.value)}>
                            {RESULTS.map((r) => (
                                <option key={r} value={r}>
                                    {r}
                                </option>
                            ))}
                        </select>
                    </div>
                    <Field name="requesterComments" label="Comentarios Solicitante" value={task.requesterComments} onChange={this.set} />
                    <Field name="evidenceLinkText" label="Anexo Evidencias (texto)" value={task.evidenceLinkText} onChange={this.set} />
                    <Field name="evidenceLinkUrl" label="Anexo Evidencias (enlace)" value={task.evidenceLinkUrl} onChange={this.set} mono />
                </div>
            </div>
        );
    }

    render() {
        const { lookup } = this.props;
        const task = lookup.task;
        const tab = this.state.tab;
        return (
            <section className="panel editor">
                {this.renderHeader()}
                {lookup.warnings.length ? (
                    <ul className="alert alert-warn warnings">
                        {lookup.warnings.map((w, i) => (
                            <li key={i}>{w}</li>
                        ))}
                    </ul>
                ) : null}
                <nav className="tabs" role="tablist">
                    {TABS.map((t) => (
                        <button
                            key={t.key}
                            role="tab"
                            aria-selected={tab === t.key}
                            className={'tab' + (tab === t.key ? ' tab-active' : '')}
                            onClick={() => this.setState({ tab: t.key })}
                        >
                            <Icon name={t.icon} /> {t.label}
                        </button>
                    ))}
                </nav>
                {tab === 'description' ? this.renderDescription(task) : null}
                {tab === 'design' ? this.renderDesign(task) : null}
                {tab === 'log' ? this.renderLog(task) : null}
                {tab === 'uat' ? this.renderUat(task) : null}
            </section>
        );
    }
}
