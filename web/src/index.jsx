import React from 'react';
import ReactDOM from 'react-dom';
import { api, onUnauthorized } from './api';
import { loadAccount, loadSession, saveAccount, saveMicrosoft, saveSession } from './storage';
import Dashboard from './Dashboard';
import Icon, { BrandLogo } from './icons';

class Login extends React.Component {
    constructor(props) {
        super(props);
        const account = loadAccount();
        this.state = {
            username: account.username || '',
            email: account.email || '',
            password: '',
            busy: false,
            error: props.notice || ''
        };
        this.submit = this.submit.bind(this);
    }

    submit(e) {
        e.preventDefault();
        if (this.state.busy) {
            return;
        }
        this.setState({ busy: true, error: '' });
        const username = this.state.username.trim();
        const email = this.state.email.trim();
        api.login(username, email, this.state.password).then(
            (data) => {
                saveAccount(username, email);
                this.setState({ busy: false, password: '' });
                this.props.onLogin({ token: data.token, user: data.user, mock: data.mock, version: data.version });
            },
            (error) => this.setState({ busy: false, error: error.message })
        );
    }

    render() {
        const { username, email, password, busy, error } = this.state;
        return (
            <div className="login-page">
                <form className="login-card" onSubmit={this.submit}>
                    <div className="login-brand">
                        <BrandLogo size={52} />
                        <div>
                            <h1>
                                GEIR
                            </h1>
                            <p>Guía de Evaluación e Implementación de Requerimientos</p>
                        </div>
                    </div>
                    <p className="muted">
                        Entra con tus credenciales de red. El usuario se usa en Azure DevOps y GitLab; el correo, en
                        OneDrive (Microsoft 365).
                    </p>
                    <label htmlFor="username">Usuario de red (GitLab / Azure DevOps)</label>
                    <div className="input-icon">
                        <Icon name="user" />
                        <input
                            id="username"
                            autoFocus={!username}
                            autoComplete="username"
                            placeholder="ej. usuario  o  DOMINIO\usuario"
                            value={username}
                            onChange={(e) => this.setState({ username: e.target.value })}
                        />
                    </div>
                    <label htmlFor="email">Correo de Microsoft (OneDrive)</label>
                    <div className="input-icon">
                        <Icon name="mail" />
                        <input
                            id="email"
                            type="email"
                            autoComplete="email"
                            placeholder="ej. nombre_apellido@empresa.com"
                            value={email}
                            onChange={(e) => this.setState({ email: e.target.value })}
                        />
                    </div>
                    <label htmlFor="password">Contraseña</label>
                    <div className="input-icon">
                        <Icon name="lock" />
                        <input
                            id="password"
                            type="password"
                            autoComplete="current-password"
                            autoFocus={!!username}
                            value={password}
                            onChange={(e) => this.setState({ password: e.target.value })}
                        />
                    </div>
                    {error ? (
                        <div className="alert alert-error alert-icon">
                            <Icon name="alert" size={18} />
                            <span>{error}</span>
                        </div>
                    ) : null}
                    <button className="btn btn-primary btn-block" type="submit" disabled={busy || !username.trim() || !email.trim() || !password}>
                        <Icon name={busy ? 'loader' : 'login'} spin={busy} />
                        {busy ? 'Validando credenciales…' : 'Iniciar sesión'}
                    </button>
                    <p className="hint">
                        <Icon name="shield" size={13} />{' '}
                        {busy
                            ? 'Validando contra Azure DevOps, puede tardar unos segundos.'
                            : 'Tu sesión se guarda solo en esta pestaña y se borra al cerrarla. Nunca se guarda la contraseña.'}
                    </p>
                </form>
            </div>
        );
    }
}

class App extends React.Component {
    constructor(props) {
        super(props);
        const stored = loadSession();
        this.state = { session: stored, checking: !!stored, notice: '' };
        this.login = this.login.bind(this);
        this.logout = this.logout.bind(this);
    }

    componentDidMount() {
        onUnauthorized((message) => this.expire(message));
        if (this.state.session) {
            // The server keeps credentials in memory, so a restart invalidates the token.
            api.me().then(
                (data) => {
                    const session = Object.assign({}, this.state.session, { user: data.user, mock: data.mock, version: data.version });
                    saveSession(session);
                    this.setState({ session: session, checking: false });
                },
                () => this.setState({ checking: false })
            );
        }
    }

    expire(message) {
        saveSession(null);
        this.setState({ session: null, checking: false, notice: message || 'Tu sesión expiró.' });
    }

    login(session) {
        saveSession(session);
        this.setState({ session: session, notice: '' });
    }

    logout() {
        api.logout().catch(() => null);
        saveSession(null);
        saveMicrosoft(null); // the Microsoft sign-in ends with the GEIR session
        this.setState({ session: null, notice: '' });
    }

    render() {
        if (this.state.checking) {
            return <div className="splash">Cargando…</div>;
        }
        if (!this.state.session) {
            return <Login onLogin={this.login} notice={this.state.notice} />;
        }
        return <Dashboard session={this.state.session} onLogout={this.logout} />;
    }
}

ReactDOM.render(<App />, document.getElementById('root'));
