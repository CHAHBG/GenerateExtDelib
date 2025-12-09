import * as XLSX from 'xlsx';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
// @ts-ignore
import DocxMerger from 'docx-merger';

interface FileSet {
    indiv: File | null;
    coll: File | null;
    coord_pi: File | null;
    coord_pc: File | null;
    tpl_indiv: File | null;
    tpl_coll: File | null;
}

export class GeneratorEngineJS {
    private logs: (msg: string) => void;

    constructor(logCallback: (msg: string) => void) {
        this.logs = logCallback;
    }

    private log(msg: string) {
        this.logs(msg);
    }

    private yieldToUI(): Promise<void> {
        return new Promise(resolve => setTimeout(resolve, 0));
    }

    async run(files: FileSet) {
        this.log("🚀 Démarrage Engine JS (Ultra-Rapide)...");
        try {
            const zip = new JSZip();
            const folderIndiv = zip.folder("Individuelles");
            const folderColl = zip.folder("Collectives");

            const dfIndiv = await this.readExcel(files.indiv, "Indiv");
            const dfColl = await this.readExcel(files.coll, "Coll");
            const dfCoordPI = await this.readExcel(files.coord_pi, "Coords PI");
            const dfCoordPC = await this.readExcel(files.coord_pc, "Coords PC");

            if (dfIndiv.length > 0) this.log(`🔍 Colonnes INDIV: ${Object.keys(dfIndiv[0]).slice(0, 5).join(', ')}...`);
            if (dfCoordPI.length > 0) this.log(`🔍 Colonnes Coords: ${Object.keys(dfCoordPI[0]).slice(0, 5).join(', ')}...`);

            const tplIndivBuffer = files.tpl_indiv ? await files.tpl_indiv.arrayBuffer() : null;
            const tplCollBuffer = files.tpl_coll ? await files.tpl_coll.arrayBuffer() : null;

            const indivBlobs: ArrayBuffer[] = [];
            const collBlobs: ArrayBuffer[] = [];

            // Generate Individual documents
            if (dfIndiv.length > 0 && tplIndivBuffer) {
                this.log(`📄 Génération: ${dfIndiv.length} Individuelles...`);
                let count = 0;
                for (const row of dfIndiv) {
                    const nicad = this.cleanID(row.nicad);
                    try {
                        const data = this.prepareData(row, nicad, dfCoordPI);
                        const blob = this.generateDoc(tplIndivBuffer, data);
                        const ab = await blob.arrayBuffer();
                        indivBlobs.push(ab);
                        folderIndiv?.file(`Extrait_PI_${nicad}.docx`, blob);
                        count++;
                        if (count % 50 === 0) {
                            this.log(`   ... ${count} faits`);
                            await this.yieldToUI();
                        }
                    } catch (e) {
                        this.log(`   ❌ Err ${nicad}: ${e}`);
                    }
                }
                this.log(`   ✓ ${count} Individuelles générées`);
            }

            // Generate Collective documents
            if (dfColl.length > 0 && tplCollBuffer) {
                this.log(`📄 Génération: ${dfColl.length} Collectives...`);
                let count = 0;
                for (const row of dfColl) {
                    const nicad = this.cleanID(row.nicad);
                    try {
                        const data = this.prepareDataColl(row, nicad, dfCoordPC);
                        const blob = this.generateDoc(tplCollBuffer, data);
                        const ab = await blob.arrayBuffer();
                        collBlobs.push(ab);
                        folderColl?.file(`Extrait_PC_${nicad}.docx`, blob);
                        count++;
                        if (count % 50 === 0) await this.yieldToUI();
                    } catch (e) {
                        this.log(`   ❌ Err ${nicad}: ${e}`);
                    }
                }
                this.log(`   ✓ ${count} Collectives générées`);
            }

            // Merge Individual documents
            if (indivBlobs.length > 0) {
                this.log(`📑 Fusion des ${indivBlobs.length} Individuelles...`);
                try {
                    const mergedIndiv = await this.mergeDocuments(indivBlobs);
                    zip.file("TOUS_LES_EXTRAITS_INDIVIDUELS.docx", mergedIndiv);
                    this.log(`   ✓ Fichier fusionné créé`);
                } catch (e) {
                    this.log(`   ⚠️ Fusion échouée: ${e}`);
                }
            }

            // Merge Collective documents
            if (collBlobs.length > 0) {
                this.log(`📑 Fusion des ${collBlobs.length} Collectives...`);
                try {
                    const mergedColl = await this.mergeDocuments(collBlobs);
                    zip.file("TOUS_LES_EXTRAITS_COLLECTIVES.docx", mergedColl);
                    this.log(`   ✓ Fichier fusionné créé`);
                } catch (e) {
                    this.log(`   ⚠️ Fusion échouée: ${e}`);
                }
            }

            this.log("📦 Compression ZIP...");
            const content = await zip.generateAsync({ type: "blob" });
            saveAs(content, "Extraits_Generes.zip");
            this.log("✅ TERMINÉ ! Téléchargement lancé.");
        } catch (e) {
            this.log(`🔥 ERREUR CRITIQUE: ${e}`);
            console.error(e);
        }
    }

    private async mergeDocuments(docs: ArrayBuffer[]): Promise<Blob> {
        const docxMerger = new DocxMerger({}, docs);
        return new Promise((resolve, _reject) => {
            docxMerger.save('blob', (data: Blob) => {
                resolve(data);
            });
        });
    }

    private async readExcel(file: File | null, name: string): Promise<any[]> {
        if (!file) return [];
        this.log(`📖 Lecture ${name}...`);
        const ab = await file.arrayBuffer();
        const wb = XLSX.read(ab);
        const ws = wb.Sheets[wb.SheetNames[0]];
        return XLSX.utils.sheet_to_json(ws);
    }

    private cleanID(val: any): string {
        if (val === undefined || val === null) return "";
        let s = String(val).trim();
        if (s.endsWith('.0')) s = s.substring(0, s.length - 2);
        return s;
    }

    private splitCoords(pts: any[]): any[] {
        if (!pts || pts.length === 0) return [];
        const mid = Math.ceil(pts.length / 2);
        const left = pts.slice(0, mid);
        const right = pts.slice(mid);
        const rows = [];
        for (let i = 0; i < left.length; i++) {
            rows.push({
                pt1: left[i].pt,
                x1: left[i].x,
                y1: left[i].y,
                pt2: right[i] ? right[i].pt : "",
                x2: right[i] ? right[i].x : "",
                y2: right[i] ? right[i].y : ""
            });
        }
        return rows;
    }

    private prepareData(row: any, nicad: string, coords: any[]) {
        const pts = coords
            .filter(c => this.cleanID(c.nicad) === nicad)
            .sort((a, b) => (a.vertex_index || 0) - (b.vertex_index || 0))
            .map((c, i) => ({
                pt: `P${i + 1}`,
                x: Number(c.X ?? c.x ?? c.x_centroid ?? 0).toFixed(2),
                y: Number(c.Y ?? c.y ?? c.y_centroid ?? 0).toFixed(2)
            }));

        return {
            nicad: nicad,
            Nom: row.Nom || "",
            Prenom: row.Prenom || "",
            Village: row.Village || "",
            superficie: row.superficie || "",
            type_usag: row.type_usag || "",
            Num_piece: row.Num_piece || "",
            Type_piece: row.Type_piece || "",
            Date_naissance: row.Date_naissance || row.date_naiss || "",
            Telephone: row.Telephone || "",
            coords: pts,
            coords_split: this.splitCoords(pts)
        };
    }

    private prepareDataColl(row: any, nicad: string, coords: any[]) {
        const pts = coords
            .filter(c => this.cleanID(c.nicad) === nicad)
            .map((c, i) => ({
                pt: `P${i + 1}`,
                x: Number(c.X ?? c.x ?? c.x_centroid ?? 0).toFixed(2),
                y: Number(c.Y ?? c.y ?? c.y_centroid ?? 0).toFixed(2)
            }));

        const rawNoms = (row.Nom || "").toString();
        const rawPrenoms = (row.Prenom || "").toString();
        const rawPieces = (row.Numero_piece || row.Num_piece || "").toString();

        const noms = rawNoms.split('\n');
        const prenoms = rawPrenoms.split('\n');
        const pieces = rawPieces.split('\n');

        const len = Math.max(noms.length, prenoms.length, pieces.length);
        const beneficiaires = [];
        for (let i = 0; i < len; i++) {
            if ((noms[i] && noms[i].trim()) || (prenoms[i] && prenoms[i].trim())) {
                beneficiaires.push({
                    Nom: noms[i]?.trim() || "",
                    Prenom: prenoms[i]?.trim() || "",
                    CNI: pieces[i]?.trim() || ""
                });
            }
        }

        return {
            nicad: nicad,
            Village: row.Village || "",
            superficie: row.superficie || "",
            type_usa: row.type_usa || "",
            beneficiaires: beneficiaires,
            coords: pts,
            coords_split: this.splitCoords(pts),
            Nom: rawNoms.replace(/\n/g, ' / '),
            Prenom: rawPrenoms.replace(/\n/g, ' / '),
            Num_piece: rawPieces.replace(/\n/g, ' / ')
        };
    }

    private generateDoc(tplBuffer: ArrayBuffer, data: any): Blob {
        const zip = new PizZip(tplBuffer);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            delimiters: { start: '«', end: '»' },
            nullGetter: () => ""
        });
        doc.render(data);
        return doc.getZip().generate({
            type: "blob",
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        });
    }
}
