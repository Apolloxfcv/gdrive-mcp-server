import ExcelJS from "exceljs";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  patchDocument,
  PatchType,
} from "docx";
import JSZip from "jszip";
import { MAX_UNCOMPRESSED_BYTES, UserFacingError } from "./tool-guards";

/**
 * Utilitaires de lecture/ecriture pour fichiers Office binaires (.xlsx, .docx)
 * stockes tels quels sur Google Drive (pas des Google Sheets/Docs natifs).
 *
 * Principe stateless : ces fonctions travaillent uniquement sur des Buffer
 * en memoire, le temps d'une requete HTTP. Rien n'est jamais ecrit sur disque
 * cote serveur. Le fichier complet est telecharge depuis Drive, modifie en
 * memoire, puis re-uploade en entier (pas d'edition incrementale possible
 * sur un binaire ZIP/XML comme le sont .xlsx et .docx).
 */

// ---------- ZIP ----------

/**
 * Protection contre les "zip bombs" : un .xlsx/.docx de quelques Mo peut se
 * decompresser en plusieurs Go et faire tomber la fonction. On lit le
 * repertoire central (tailles declarees) avant toute decompression.
 */
async function loadZipSafely(buffer: Buffer): Promise<JSZip> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    throw new UserFacingError("Fichier Office invalide (archive ZIP illisible).");
  }
  let total = 0;
  let entries = 0;
  zip.forEach((_path, file) => {
    entries++;
    total += (file as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0;
  });
  if (entries > 10_000 || total > MAX_UNCOMPRESSED_BYTES) {
    throw new UserFacingError("Fichier Office refuse : contenu decompresse trop volumineux.");
  }
  return zip;
}

// ---------- XLSX ----------

export async function readXlsxAsJson(buffer: Buffer) {
  await loadZipSafely(buffer);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer as any);

  const sheets: Record<string, (string | number | boolean | null)[][]> = {};
  workbook.eachSheet((worksheet) => {
    const rows: (string | number | boolean | null)[][] = [];
    worksheet.eachRow({ includeEmpty: true }, (row) => {
      const values = (row.values as any[]).slice(1);
      rows.push(values.map((v) => (v === undefined ? null : v)));
    });
    sheets[worksheet.name] = rows;
  });
  return sheets;
}

export async function updateXlsxCells(
  buffer: Buffer,
  sheetName: string,
  updates: { cell: string; value: string | number | boolean }[]
): Promise<Buffer> {
  await loadZipSafely(buffer);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer as any);

  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    throw new UserFacingError(
      `Feuille "${sheetName}" introuvable. Feuilles disponibles: ${workbook.worksheets
        .map((w) => w.name)
        .join(", ")}`
    );
  }

  for (const { cell, value } of updates) {
    worksheet.getCell(cell).value = value;
  }

  const out = await workbook.xlsx.writeBuffer();
  return Buffer.from(out);
}

export async function createXlsx(
  sheetName: string,
  rows: (string | number | boolean | null)[][]
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(sheetName);
  rows.forEach((row) => worksheet.addRow(row));
  const out = await workbook.xlsx.writeBuffer();
  return Buffer.from(out);
}

// ---------- DOCX ----------

export async function readDocxAsText(buffer: Buffer): Promise<string> {
  const zip = await loadZipSafely(buffer);
  const documentXml = await zip.file("word/document.xml")?.async("string");
  if (!documentXml) throw new UserFacingError("document.xml introuvable dans ce .docx");

  const text = documentXml
    .replace(/<w:p[ >]/g, "\n<w:p>")
    .replace(/<[^>]+>/g, "")
    .replace(/\n{2,}/g, "\n")
    .trim();
  return text;
}

export async function createDocx(paragraphs: string[]): Promise<Buffer> {
  const doc = new Document({
    sections: [
      {
        children: paragraphs.map(
          (text) => new Paragraph({ children: [new TextRun(text)] })
        ),
      },
    ],
  });
  const out = await Packer.toBuffer(doc);
  return Buffer.from(out);
}

export async function patchDocxPlaceholders(
  buffer: Buffer,
  replacements: Record<string, string>
): Promise<Buffer> {
  await loadZipSafely(buffer);
  const patches: Record<string, any> = {};
  for (const [key, value] of Object.entries(replacements)) {
    patches[key] = {
      type: PatchType.PARAGRAPH,
      children: [new TextRun(value)],
    };
  }

  // "nodebuffer" est la valeur valide pour un environnement Node.js/serverless
  // (pas "buffer", qui n'existe pas dans le type OutputByType de la lib docx).
  const patched = await patchDocument({
    outputType: "nodebuffer",
    data: buffer,
    patches,
  });
  return Buffer.from(patched);
}
