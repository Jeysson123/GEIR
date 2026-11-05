// Regenerates templates/N - GEIR.xlsx (used when the OneDrive folder has no GEIR yet)
// from a filled GEIR: keeps every sheet's format and leaves one empty slot per section.
//
//   npm run template -- "C:\Users\me\Downloads\N - GEIR.xlsx"
import * as fs from 'fs';
import * as path from 'path';
import { GeirDocument } from '../src/domain/document';

async function main() {
    const source = process.argv[2];
    if (!source) {
        throw new Error('Uso: npm run template -- "<ruta a un N - GEIR.xlsx>"');
    }
    const doc = await GeirDocument.load(fs.readFileSync(source));
    await doc.clearAllTasks();
    const target = path.resolve(__dirname, '..', 'templates', 'N - GEIR.xlsx');
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, await doc.save());
    console.log(`Plantilla creada: ${target}`);
}

main().catch((error) => {
    console.error(error.message ?? error);
    process.exit(1);
});
