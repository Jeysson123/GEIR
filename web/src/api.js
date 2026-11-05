import { loadMicrosoft, loadSession, saveMicrosoft } from './storage';

var unauthorizedHandler = null;

export function onUnauthorized(handler) {
    unauthorizedHandler = handler;
}

function request(method, url, body) {
    var session = loadSession();
    var headers = { 'Content-Type': 'application/json' };
    if (session && session.token) {
        headers.Authorization = 'Bearer ' + session.token;
    }
    return fetch(url, {
        method: method,
        headers: headers,
        body: body ? JSON.stringify(body) : undefined
    }).then(function (res) {
        keepMicrosoft(res);
        return res
            .json()
            .catch(function () {
                return {};
            })
            .then(function (data) {
                if (res.status === 401 && url !== '/api/auth/login' && unauthorizedHandler) {
                    unauthorizedHandler(data.error);
                }
                if (!res.ok) {
                    var error = new Error(data.error || 'Error ' + res.status);
                    error.status = res.status;
                    throw error;
                }
                return data;
            });
    }, function () {
        throw new Error('No se pudo contactar el servidor GEIR. ¿Sigue corriendo "npm start"?');
    });
}

// The server sends a new Microsoft token of this user in this header; it is kept only in
// this tab's sessionStorage (empty value: forget it).
function keepMicrosoft(res) {
    var token = res.headers.get('X-GEIR-Microsoft');
    if (token === null) {
        return;
    }
    var session = loadSession();
    var email = session && session.user && session.user.email;
    saveMicrosoft(token && email ? { email: email, refreshToken: token } : null);
}

export var api = {
    login: function (username, email, password) {
        return request('POST', '/api/auth/login', {
            username: username,
            email: email,
            password: password,
            microsoft: loadMicrosoft()
        });
    },
    me: function () {
        return request('GET', '/api/auth/me');
    },
    logout: function () {
        return request('POST', '/api/auth/logout');
    },
    microsoftStatus: function () {
        return request('GET', '/api/microsoft/status');
    },
    openDocument: function (location) {
        return request('POST', '/api/document/open', { location: location });
    },
    lookupTask: function (location, taskId, refresh) {
        return request('POST', '/api/tasks/lookup', { location: location, taskId: taskId, refresh: !!refresh });
    },
    saveTask: function (location, task, evidenceChanged) {
        return request('POST', '/api/tasks/save', { location: location, task: task, evidenceChanged: evidenceChanged });
    },
    deleteTask: function (location, taskId) {
        return request('POST', '/api/tasks/delete', { location: location, taskId: taskId });
    },
    adminGenerate: function (payload) {
        return request('POST', '/api/admin/generate', payload);
    },
    // Compiles the branch (git pull + Maven / npm run buildprod) into the Deploy folder.
    adminBuild: function (payload) {
        return request('POST', '/api/admin/build', payload);
    },
    adminBuildDefaults: function () {
        return request('GET', '/api/admin/build-defaults');
    },
    adminBuildStatus: function (id) {
        return request('GET', '/api/admin/build/' + encodeURIComponent(id));
    },
    // Downloads a generated file (sent with the session token) and saves it with its name.
    adminDownload: function (name) {
        var session = loadSession();
        return fetch('/api/admin/files/' + encodeURIComponent(name), {
            headers: { Authorization: 'Bearer ' + (session && session.token) }
        }).then(function (res) {
            if (!res.ok) {
                return res.json().then(
                    function (data) {
                        throw new Error(data.error || 'Error ' + res.status);
                    },
                    function () {
                        throw new Error('Error ' + res.status);
                    }
                );
            }
            return res.blob().then(function (blob) {
                var url = URL.createObjectURL(blob);
                var link = document.createElement('a');
                link.href = url;
                link.download = name;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                setTimeout(function () {
                    URL.revokeObjectURL(url);
                }, 1000);
            });
        });
    }
};
