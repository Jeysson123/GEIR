import * as zlib from 'zlib';
import type { WorkItemInfo, WorkItemSummary } from '../infrastructure/azure-devops/azure-devops.client';
import type { TaskChanges } from '../infrastructure/gitlab/gitlab.client';
import { Environment } from '../infrastructure/config/environment';

// Fixture data for GEIR_MOCK=true: lets the UI and the Excel flow run without VPN.

/** A plain PNG (header bar + content blocks) standing in for a test screenshot. */
function evidencePng(width = 960, height = 540): Buffer {
    const raw = Buffer.alloc((width * 3 + 1) * height);
    for (let y = 0; y < height; y++) {
        const row = y * (width * 3 + 1);
        raw[row] = 0;
        for (let x = 0; x < width; x++) {
            let rgb = [245, 247, 250];
            if (y < 56) rgb = [0, 70, 90];
            else if (x > 40 && x < 300 && y > 90 && y < 500) rgb = [225, 232, 240];
            else if (x > 340 && x < 920 && y > 90 && y < 150) rgb = [46, 134, 222];
            else if (x > 340 && x < 920 && y > 180 && y < 500 && (y - 180) % 40 < 28) rgb = [255, 255, 255];
            raw.set(rgb, row + 1 + x * 3);
        }
    }
    const chunk = (type: string, data: Buffer) => {
        const len = Buffer.alloc(4);
        len.writeUInt32BE(data.length);
        const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
        const crc = Buffer.alloc(4);
        crc.writeUInt32BE(crc32(body) >>> 0);
        return Buffer.concat([len, body, crc]);
    };
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(width, 0);
    ihdr.writeUInt32BE(height, 4);
    ihdr[8] = 8; // bit depth
    ihdr[9] = 2; // RGB
    return Buffer.concat([
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        chunk('IHDR', ihdr),
        chunk('IDAT', zlib.deflateSync(raw)),
        chunk('IEND', Buffer.alloc(0))
    ]);
}

function crc32(buf: Buffer): number {
    let c = ~0;
    for (const byte of buf) {
        c ^= byte;
        for (let k = 0; k < 8; k++) {
            c = c & 1 ? (c >>> 1) ^ 0xedb88320 : c >>> 1;
        }
    }
    return ~c;
}

export function mockWorkItem(id: string): WorkItemInfo {
    const base = `${Environment.services.azureDevOpsUrl}/${Environment.geir.azure.collection}/${Environment.geir.azure.project}`;
    return {
        id,
        title: `Módulo Demo: Reporte de pedidos pendientes (demo ${id})`,
        workItemType: 'Product Backlog Item',
        state: 'Done',
        descriptionHtml:
            '<div>Como administrador quiero ver un reporte de los pedidos pendientes de los últimos 30 días ' +
            'para dar seguimiento a los clientes y priorizar su atención.</div>' +
            '<div>El reporte debe permitir filtrar por cliente y exportarse a Excel.</div>',
        acceptanceCriteriaHtml:
            '<div>Dado que el administrador inició sesión en el sistema y está en Inicio &gt; Reportes</div>' +
            '<div>Cuando abre el reporte "Pedidos pendientes" y filtra por cliente</div>' +
            '<div>Entonces ve los pedidos pendientes del cliente con su fecha y estado, y puede exportarlos</div>',
        iterationPath: `${Environment.geir.azure.project}\\Equipo Sprint 20 - 2026`,
        editUrl: `${base}/_workitems/edit/${id}`,
        taskboardUrl: `${base}/_sprints/taskboard/Equipo/Proyecto/Equipo%20Sprint%2020%20-%202026?workitem=${id}`,
        taskboardText: 'Equipo Sprint 20 - 2026',
        discussionImages: [
            {
                url: `${base}/_apis/wit/attachments/demo?fileName=evidencia.png`,
                base64: evidencePng().toString('base64'),
                extension: 'png',
                author: 'Usuario Demo',
                date: new Date().toISOString(),
                evidenceHint: true
            }
        ],
        warnings: ['Modo demostración (GEIR_MOCK=true): datos de ejemplo, no se consultó Azure DevOps.']
    };
}

export function mockChanges(id: string): TaskChanges {
    const project = Environment.geir.gitlab.projects[0] ?? 'grupo/proyecto';
    return {
        sources: [
            {
                kind: 'merge_request',
                project,
                ref: '!412',
                title: `${id} - Reporte de pedidos pendientes`,
                url: `${Environment.services.gitlabUrl}/${project}/-/merge_requests/412`
            }
        ],
        files: [
            {
                project,
                path: 'backend/src/main/java/com/example/app/report/PendingOrdersService.java',
                oldPath: 'backend/src/main/java/com/example/app/report/PendingOrdersService.java',
                status: 'added',
                diff: [
                    '@@ -0,0 +1,40 @@',
                    '+public class PendingOrdersService {',
                    '+    public List<PendingOrder> findByCustomer(String customerId, int days) {',
                    '+        return repository.findPending(customerId, LocalDate.now().minusDays(days));',
                    '+    }',
                    '+    public byte[] exportToExcel(List<PendingOrder> cases) {'
                ].join('\n')
            },
            {
                project,
                path: 'backend/src/main/java/com/example/app/rest/ReportREST.java',
                oldPath: 'backend/src/main/java/com/example/app/rest/ReportREST.java',
                status: 'modified',
                diff: [
                    '@@ -88,6 +88,18 @@ public class ReportREST {',
                    '+    @GET',
                    '+    @Path("/report/pending-orders")',
                    '+    public Response getPendingOrders(@QueryParam("customer") String customer) {',
                    '@@ -140,7 +152,7 @@ public Response getReports(String type) {',
                    '-        return Response.ok(reports).build();',
                    '+        return Response.ok(withPending(reports)).build();'
                ].join('\n')
            },
            {
                project,
                path: 'frontend/src/Report/PendingOrdersReport.js',
                oldPath: 'frontend/src/Report/PendingOrdersReport.js',
                status: 'added',
                diff: [
                    '@@ -0,0 +1,120 @@',
                    '+class PendingOrdersReport extends React.Component {',
                    '+  handleFilter(customer) {',
                    "+    fetchJson('/api/report/pending-orders?customer=' + customer)",
                    '+  }',
                    '+  handleExport() {'
                ].join('\n')
            },
            {
                project,
                path: 'frontend/src/Report/ReportTabs.js',
                oldPath: 'frontend/src/Report/ReportTabs.js',
                status: 'modified',
                diff: ['@@ -12,6 +12,7 @@ function ReportTabs(props) {', "+    { key: 'pending', label: 'Pedidos pendientes' },"].join('\n')
            }
        ],
        warnings: ['Modo demostración (GEIR_MOCK=true): cambios de ejemplo, no se consultó GitLab.']
    };
}

/**
 * Demo summary for the admin options: a user story whose dates spread over the sprint
 * (by ID, so they are stable), and every fifth ID a bug.
 */
export function mockSummary(id: string): WorkItemSummary {
    const n = Number(id) || 0;
    const created = new Date(Date.UTC(2026, 3, 1 + (n % 45)));
    return {
        id,
        title: `Historia de demostración ${id}`,
        workItemType: n % 5 === 0 ? 'Bug' : 'User Story',
        state: 'Active',
        createdDate: created.toISOString().slice(0, 10),
        startDate: '',
        targetDate: '',
        closedDate: '',
        parentId: '',
        assignedTo: 'Usuario Demo',
        editUrl: `${Environment.services.azureDevOpsUrl}/${Environment.geir.azure.collection}/${Environment.geir.azure.project}/_workitems/edit/${id}`
    };
}
