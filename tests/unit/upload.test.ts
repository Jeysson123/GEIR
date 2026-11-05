import { test } from 'node:test';
import assert from 'node:assert/strict';
import { OneDriveFilesClient, UPLOAD_CHUNK } from '../../src/infrastructure/microsoft/onedrive-files.client';
import { RepoView, frontScript, projectOf } from '../../src/domain/build/projects';

/** A SharePoint that keeps files in memory and understands the REST calls GEIR makes. */
function fakeSharePoint() {
    const files = new Map<string, Buffer>();
    const pending = new Map<string, Buffer[]>();
    const calls: string[] = [];
    const response = (status: number, json: any = {}) => ({
        status: () => status,
        headers: () => ({}),
        text: async () => JSON.stringify(json),
        body: async () => Buffer.alloc(0)
    });
    const decoded = (url: string, fn: string) => decodeURIComponent(new RegExp(`${fn}\\(decodedurl='([^']+)'`).exec(url)?.[1] ?? '');
    const http = {
        fetch: async (url: string, options: any) => {
            const api = url.replace(/^https:\/\/[^/]+/, '');
            const action = [...api.matchAll(/\/(StartUpload|ContinueUpload|FinishUpload|CancelUpload|AddUsingPath)\(/g)].pop()?.[1] ?? 'GET';
            calls.push(`${options.method} ${action} ${options.data?.length ?? 0}`);
            if (/GetFolderByServerRelativePath/.test(api) && options.method === 'GET') {
                return response(200, { Exists: true });
            }
            if (/AddUsingPath/.test(api)) {
                const folder = decoded(api, 'GetFolderByServerRelativePath');
                const name = decoded(api, 'AddUsingPath');
                files.set(`${folder}/${name}`, Buffer.from(options.data ?? ''));
                return response(200);
            }
            const file = decoded(api, 'GetFileByServerRelativePath');
            const id = /uploadId=guid'([^']+)'/.exec(api)?.[1] ?? '';
            if (/StartUpload/.test(api)) {
                pending.set(id, [Buffer.from(options.data)]);
                return response(200);
            }
            if (/ContinueUpload/.test(api)) {
                pending.get(id)!.push(Buffer.from(options.data));
                return response(200);
            }
            if (/FinishUpload/.test(api)) {
                files.set(file, Buffer.concat([...pending.get(id)!, Buffer.from(options.data)]));
                return response(200);
            }
            if (options.method === 'GET' && file) {
                const data = files.get(file);
                return response(200, data ? { Exists: true, Length: String(data.length), UniqueId: 'x' } : { Exists: false });
            }
            return response(404);
        }
    };
    const client = new OneDriveFilesClient(async () => ({ http: http as any, bearer: 't' }));
    return { client, files, calls };
}

const site = 'https://contoso-my.sharepoint.com/personal/u';
const at = (p: string) => ({ kind: 'sharepoint' as const, siteUrl: site, serverRelativePath: `/personal/u/Documents/Sprint/Deploy/${p}` });

test('un archivo pequeño (front) se sube en una sola petición y se verifica su tamaño', async () => {
    const sp = fakeSharePoint();
    const bundle = Buffer.from('/* bundle */');
    await sp.client.write(at('frontend/dist/js/bundle.js'), bundle);
    assert.equal(await sp.client.size(at('frontend/dist/js/bundle.js')), bundle.length);
    assert.equal(sp.calls.filter((c) => /StartUpload|FinishUpload/.test(c)).length, 0);
});

test('un jar grande se sube por partes y queda completo en OneDrive', async () => {
    const sp = fakeSharePoint();
    const jar = Buffer.alloc(UPLOAD_CHUNK * 2 + 12345, 7);
    jar[jar.length - 1] = 9;
    await sp.client.write(at('backend/jar_interfaces/billing.jar'), jar);
    const saved = sp.files.get('/personal/u/Documents/Sprint/Deploy/backend/jar_interfaces/billing.jar')!;
    assert.ok(saved.equals(jar), 'el contenido es idéntico');
    assert.deepEqual(
        sp.calls.filter((c) => /Upload/.test(c)).map((c) => c.split(' ')[1]),
        ['StartUpload', 'ContinueUpload', 'FinishUpload']
    );
    assert.ok(sp.calls.every((c) => Number(c.split(' ')[2]) <= UPLOAD_CHUNK), 'ninguna petición pasa de 8 MB');
    assert.equal(await sp.client.size(at('backend/jar_interfaces/billing.jar')), jar.length);
    assert.equal(await sp.client.size(at('no-existe.jar')), null);
});

test('una carpeta con package.json que tiene buildprod es el front aunque tenga pom.xml', () => {
    const files: Record<string, string> = {
        'web/pom.xml': '<project><artifactId>web</artifactId></project>',
        'web/package.json': '{"scripts":{"buildprod":"webpack -p"}}',
        'api/pom.xml': '<project><artifactId>api</artifactId></project>',
        'api/tools/package.json': '{"scripts":{"lint":"eslint"}}'
    };
    const repo: RepoView = { exists: (r) => r in files, read: (r) => files[r] ?? null, poms: () => [] };
    assert.equal(frontScript('npm run buildprod'), 'buildprod');
    assert.deepEqual(projectOf('web/src/app.jsx', repo), { dir: 'web', kind: 'front' });
    assert.deepEqual(projectOf('api/src/A.java', repo), { dir: 'api', kind: 'maven' });
    assert.deepEqual(projectOf('api/tools/x.js', repo), { dir: 'api', kind: 'maven' }, 'package.json sin buildprod no es el front');
});
