import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
    DEFAULT_BUILD,
    RepoView,
    deployRoot,
    expectedJar,
    mentionsAny,
    parsePom,
    pickJar,
    planBuild,
    readmeCommand,
    validBranch
} from '../../src/domain/build/projects';

const pom = (g: string, a: string, deps: string[] = [], extra = '') => `<?xml version="1.0"?>
<project>
  <modelVersion>4.0.0</modelVersion>
  <groupId>${g}</groupId>
  <artifactId>${a}</artifactId>
  <version>4.9</version>
  ${extra}
  <dependencies>
    ${deps.map((d) => `<dependency><groupId>${d.split(':')[0]}</groupId><artifactId>${d.split(':')[1]}</artifactId><version>4.9</version></dependency>`).join('\n')}
  </dependencies>
  <build><plugins><plugin><groupId>org.apache.maven.plugins</groupId><artifactId>maven-assembly-plugin</artifactId>
    <configuration><descriptorRefs><descriptorRef>jar-with-dependencies</descriptorRef></descriptorRefs></configuration></plugin></plugins></build>
</project>`;

// Layout of a typical repository: a core library, the API, interfaces and the front end.
const files: Record<string, string> = {
    'app-core/pom.xml': pom('com.example.app', 'app-core'),
    'app-api/pom.xml': pom('com.example.app', 'app-api', ['com.example.app:app-core']),
    'interfaces/billing/pom.xml': pom('com.example.app.interfaces.billing', 'app-interfaces-billing', ['com.example.app:app-core']),
    'interfaces/gateway/pom.xml': pom('com.example.app.interfaces', 'app-interfaces-gateway', ['com.example.app:app-core']),
    'front/package.json': '{"scripts":{"buildprod":"webpack -p"}}'
};
const repo: RepoView = {
    exists: (rel) => rel in files,
    read: (rel) => files[rel] ?? null,
    poms: () => Object.keys(files).filter((f) => f.endsWith('pom.xml'))
};
const all = () => true;

test('compila solo los proyectos que tocan los commits y los manda a su carpeta de Deploy', () => {
    const plan = planBuild(['interfaces/billing/src/main/java/A.java', 'front/src/app.jsx', 'README.md'], repo, DEFAULT_BUILD, all);
    assert.deepEqual(
        plan.map((p) => [p.name, p.reason, p.destination]),
        [
            ['app-interfaces-billing', 'cambio', 'backend/jar_interfaces'],
            ['front', 'cambio', 'frontend/dist/js']
        ]
    );
});

test('por defecto un cambio en app-core solo instala el core: no recompila todo el backend', () => {
    const plan = planBuild(['app-core/src/Core.java', 'interfaces/billing/src/A.java'], repo, DEFAULT_BUILD, all);
    assert.deepEqual(plan.map((p) => [p.name, p.reason]), [
        ['app-core', 'cambio'],
        ['app-interfaces-billing', 'cambio']
    ]);
});

test('con recompilarDependientes, un cambio en app-core recompila también lo que lo usa', () => {
    const plan = planBuild(['app-core/src/Core.java'], repo, { ...DEFAULT_BUILD, recompilarDependientes: true }, all);
    assert.equal(plan[0].name, 'app-core');
    assert.equal(plan[0].destination, '', 'el core solo se instala, no va a Deploy');
    assert.deepEqual(
        plan.slice(1).map((p) => [p.name, p.reason]).sort(),
        [
            ['app-api', 'depende'],
            ['app-interfaces-billing', 'depende'],
            ['app-interfaces-gateway', 'depende']
        ]
    );
    assert.equal(plan.find((p) => p.name === 'app-api')!.destination, 'backend');
});

test('si el servidor no tiene el core en Maven, lo compila antes', () => {
    const plan = planBuild(['app-api/src/Api.java'], repo, DEFAULT_BUILD, (p) => p.artifactId !== 'app-core');
    assert.deepEqual(plan.map((p) => [p.name, p.reason]), [
        ['app-core', 'falta'],
        ['app-api', 'cambio']
    ]);
});

test('lee groupId y versión heredados del <parent> y las dependencias reales', () => {
    const p = parsePom(`<project><parent><groupId>com.x</groupId><artifactId>p</artifactId><version>2.0</version></parent>
        <artifactId>hijo</artifactId><!-- <artifactId>comentado</artifactId> -->
        <dependencyManagement><dependencies><dependency><groupId>io.dw</groupId><artifactId>bom</artifactId></dependency></dependencies></dependencyManagement>
        <dependencies><dependency><groupId>com.x</groupId><artifactId>core</artifactId></dependency></dependencies></project>`);
    assert.deepEqual(p, {
        groupId: 'com.x',
        artifactId: 'hijo',
        version: '2.0',
        packaging: 'jar',
        dependencies: ['com.x:core'],
        packagedAs: 'jar',
        finalName: 'hijo-2.0'
    });
});

test('el pom.xml decide el jar: assembly jar-with-dependencies, shade o jar normal', () => {
    const billing = parsePom(files['interfaces/billing/pom.xml']);
    assert.equal(billing.packagedAs, 'with-dependencies');
    assert.equal(
        pickJar(['app-interfaces-billing-4.9.jar', 'app-interfaces-billing-4.9-jar-with-dependencies.jar', 'classes'], billing),
        'app-interfaces-billing-4.9-jar-with-dependencies.jar'
    );
    assert.equal(pickJar(['app-interfaces-billing-4.9.jar'], billing), null, 'sin el jar con dependencias no se toma el normal');

    // app-api: shade (the real pom), the main jar already has the dependencies.
    const api = parsePom(`<project><groupId>com.example.app</groupId><artifactId>app-api</artifactId><version>1.0-SNAPSHOT</version>
        <build><plugins><plugin><artifactId>maven-shade-plugin</artifactId><version>3.2.1</version></plugin>
        <plugin><artifactId>maven-source-plugin</artifactId></plugin></plugins></build></project>`);
    assert.equal(api.packagedAs, 'shade');
    assert.equal(
        pickJar(['original-app-api-1.0-SNAPSHOT.jar', 'app-api-1.0-SNAPSHOT-sources.jar', 'app-api-1.0-SNAPSHOT.jar'], api),
        'app-api-1.0-SNAPSHOT.jar'
    );

    // Assembly commented out, finalName with properties: plain jar with that name.
    const plain = parsePom(`<project><groupId>g</groupId><artifactId>svc</artifactId><version>2.1</version>
        <properties><env>qa</env></properties>
        <build><finalName>\${project.artifactId}-\${env}</finalName><plugins>
        <!-- <plugin><artifactId>maven-assembly-plugin</artifactId><configuration><descriptorRefs><descriptorRef>jar-with-dependencies</descriptorRef></descriptorRefs></configuration></plugin> -->
        </plugins></build></project>`);
    assert.equal(plain.packagedAs, 'jar');
    assert.equal(expectedJar(plain), 'svc-qa.jar');
});

test('usa el comando del README del proyecto y lo cambia a install si es una librería', () => {
    const readme = '# Core\nPara compilar:\n```\nmvn clean package -Dmaven.test.skip=true\n```\nmvn test';
    assert.equal(readmeCommand(readme, false, 'app-core', 'app-core'), 'mvn clean package -Dmaven.test.skip=true');
    assert.equal(readmeCommand(readme, true, 'app-core', 'app-core'), 'mvn clean install -Dmaven.test.skip=true');
    assert.equal(readmeCommand('$ mvn package && rm -rf /', false), null, 'nada de comandos encadenados');
    assert.equal(readmeCommand(null, false), null);
});

test('del README de la raíz solo toma la línea de ese proyecto, nunca un mvn que compile todo', () => {
    const root = [
        '## Compilar',
        '```',
        'mvn clean install',
        'cd app-core && mvn clean install -DskipTests',
        'cd interfaces/billing && mvn clean package -DskipTests',
        '```'
    ].join('\n');
    assert.equal(readmeCommand(root, false, '', 'interfaces/billing'), 'mvn clean package -DskipTests');
    assert.equal(readmeCommand(root, false, '', 'app-core'), 'mvn clean install -DskipTests');
    assert.equal(readmeCommand(root, false, '', 'app-api'), null, 'sin línea propia: el comando por defecto');
    assert.equal(readmeCommand('- `cd billing && mvn package`', false, 'interfaces', 'interfaces/billing'), 'mvn package');
});

test('Deploy va junto al GEIR, o junto a su carpeta GEIR', () => {
    assert.equal(deployRoot('/personal/u/Documents/CRQ12345 - Sprint - 1 - 2026/GEIR'), '/personal/u/Documents/CRQ12345 - Sprint - 1 - 2026/Deploy');
    assert.equal(deployRoot('/personal/u/Documents/Sprint 17'), '/personal/u/Documents/Sprint 17/Deploy');
    assert.equal(deployRoot('C:\\dev\\Sprint\\GEIR'), 'C:\\dev\\Sprint\\Deploy');
});

test('un commit es de la GEIR solo si su mensaje nombra el ID completo', () => {
    const ids = ['100008', '100006'];
    assert.deepEqual(mentionsAny('100008 Nuevo reporte', ids), ['100008']);
    assert.deepEqual(mentionsAny('fix #100006: pantalla\n\nRelated 100008', ids), ['100008', '100006']);
    assert.deepEqual(mentionsAny('1000080 otra historia', ids), []);
    assert.deepEqual(mentionsAny('T2-100008-pantalla', ids), ['100008']);
    assert.deepEqual(mentionsAny('Merge de la rama dev', ids), []);
});

test('solo acepta nombres de rama válidos', () => {
    assert.ok(validBranch('release/CRQ12345-sprint-1'));
    assert.ok(validBranch('feature/100006_pantalla'));
    assert.ok(!validBranch('--upload-pack=x'));
    assert.ok(!validBranch('a..b'));
    assert.ok(!validBranch('rama con espacios'));
    assert.ok(!validBranch('x;rm -rf'));
});
