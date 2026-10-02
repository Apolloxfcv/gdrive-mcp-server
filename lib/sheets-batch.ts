import { z } from "zod";
import type { sheets_v4 } from "googleapis";
import { A1_RANGE_PATTERN, columnLetters, columnNumber, parseA1Range } from "./a1";
import { UserFacingError } from "./tool-guards";

/**
 * Edition d'un Google Sheets natif par UN SEUL spreadsheets.batchUpdate.
 *
 * Pourquoi : cet appel est atomique (toutes les operations passent ou aucune)
 * et ne coute qu'une seule requete d'ecriture (quota ~60/min/utilisateur). Les
 * modifications se font en place : jamais de suppression/recreation de feuille
 * ou de fichier pour "modifier" quelque chose.
 */

const cellValue = z.union([z.string().max(50_000), z.number().finite(), z.boolean(), z.null()]);
const grid = z.array(z.array(cellValue).max(200)).min(1).max(5000);
const sheetName = z.string().min(1).max(100).describe("Nom exact de l'onglet (voir drive_read_spreadsheet)");
const a1 = z.string().regex(A1_RANGE_PATTERN, "Plage A1 invalide").describe("Plage A1 sans nom d'onglet : 'B3', 'B3:D9', 'A:C', '2:10'");
const hex = z.string().regex(/^#[0-9a-f]{6}$/i, "Couleur hex #RRGGBB attendue");

export const sheetsOpSchema = z.discriminatedUnion("op", [
  z.object({
    op: z.literal("write_values"),
    sheetName,
    startCell: z
      .string()
      .regex(/^[A-Z]{1,3}[1-9][0-9]{0,6}$/i, "Cellule de depart invalide")
      .describe("Coin haut-gauche ou ecrire, ex: 'B3'"),
    values: grid.describe("Tableau 2D de valeurs ; null = vider la cellule"),
    formulas: z
      .boolean()
      .default(true)
      .describe("true : une chaine commencant par '=' est une formule ; false : tout est du texte litteral"),
  }),
  z.object({
    op: z.literal("append_rows"),
    sheetName,
    rows: grid.describe("Lignes ajoutees apres la derniere ligne contenant des donnees"),
    formulas: z.boolean().default(true),
  }),
  z.object({ op: z.literal("clear_range"), sheetName, range: a1.describe("Plage dont les VALEURS sont videes (mise en forme conservee)") }),
  z.object({
    op: z.literal("format_range"),
    sheetName,
    range: a1,
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    fontSize: z.number().int().min(1).max(400).optional(),
    textColor: hex.optional(),
    backgroundColor: hex.optional(),
    horizontalAlignment: z.enum(["LEFT", "CENTER", "RIGHT"]).optional(),
    wrapText: z.boolean().optional(),
    numberFormat: z
      .object({
        type: z.enum(["TEXT", "NUMBER", "PERCENT", "CURRENCY", "DATE", "TIME", "DATE_TIME", "SCIENTIFIC"]),
        pattern: z.string().max(100).optional().describe("ex: '#,##0.00', 'dd/mm/yyyy'"),
      })
      .optional(),
  }),
  z.object({
    op: z.literal("add_sheet"),
    sheetName: sheetName.describe("Nom du nouvel onglet"),
    rows: z.number().int().min(1).max(100_000).default(1000),
    columns: z.number().int().min(1).max(702).default(26),
  }),
  z.object({ op: z.literal("rename_sheet"), sheetName, newName: z.string().min(1).max(100) }),
  z.object({ op: z.literal("delete_sheet"), sheetName }),
  z.object({
    op: z.literal("insert_rows"),
    sheetName,
    beforeRow: z.number().int().min(1).max(10_000_000).describe("Numero de ligne (1-based) avant laquelle inserer"),
    count: z.number().int().min(1).max(10_000).default(1),
  }),
  z.object({
    op: z.literal("delete_rows"),
    sheetName,
    startRow: z.number().int().min(1).max(10_000_000),
    count: z.number().int().min(1).max(10_000).default(1),
  }),
  z.object({
    op: z.literal("insert_columns"),
    sheetName,
    beforeColumn: z.string().regex(/^[A-Z]{1,3}$/i).describe("Lettre de colonne avant laquelle inserer, ex: 'C'"),
    count: z.number().int().min(1).max(702).default(1),
  }),
  z.object({
    op: z.literal("delete_columns"),
    sheetName,
    startColumn: z.string().regex(/^[A-Z]{1,3}$/i),
    count: z.number().int().min(1).max(702).default(1),
  }),
]);

export type SheetsOp = z.infer<typeof sheetsOpSchema>;

export type SheetMeta = { sheetId: number; title: string; rowCount: number; columnCount: number };

const MAX_TOTAL_CELLS = 100_000;

function toCellData(v: z.infer<typeof cellValue>, formulas: boolean): sheets_v4.Schema$CellData {
  if (v === null) return {};
  if (typeof v === "number") return { userEnteredValue: { numberValue: v } };
  if (typeof v === "boolean") return { userEnteredValue: { boolValue: v } };
  if (formulas && v.startsWith("=") && v.length > 1) return { userEnteredValue: { formulaValue: v } };
  return { userEnteredValue: { stringValue: v } };
}

function toRows(values: z.infer<typeof grid>, formulas: boolean): sheets_v4.Schema$RowData[] {
  return values.map((row) => ({ values: row.map((v) => toCellData(v, formulas)) }));
}

function rgb(h: string) {
  const n = parseInt(h.slice(1), 16);
  return { red: ((n >> 16) & 255) / 255, green: ((n >> 8) & 255) / 255, blue: (n & 255) / 255 };
}

function gridRange(sheetId: number, range: string): sheets_v4.Schema$GridRange {
  const r = parseA1Range(range);
  const out: sheets_v4.Schema$GridRange = { sheetId };
  if (r.startRow !== undefined) out.startRowIndex = r.startRow - 1;
  if (r.endRow !== undefined) out.endRowIndex = r.endRow;
  if (r.startCol !== undefined) out.startColumnIndex = r.startCol - 1;
  if (r.endCol !== undefined) out.endColumnIndex = r.endCol;
  return out;
}

/**
 * Traduit la liste d'operations en requetes batchUpdate. Les onglets sont
 * resolus par nom dans un modele local mis a jour operation apres operation
 * (un onglet cree/renomme/supprime plus haut dans la liste est donc pris en compte).
 */
export function buildSheetsBatch(ops: SheetsOp[], existing: SheetMeta[]) {
  const model = new Map<string, SheetMeta>(existing.map((s) => [s.title, { ...s }]));
  let nextId = existing.reduce((m, s) => Math.max(m, s.sheetId), 0) + 1;
  const requests: sheets_v4.Schema$Request[] = [];
  const plan: string[] = [];
  let totalCells = 0;

  const sheet = (name: string): SheetMeta => {
    const s = model.get(name);
    if (!s) {
      throw new UserFacingError(
        `Onglet "${name}" introuvable. Onglets disponibles: ${[...model.keys()].join(", ")}`
      );
    }
    return s;
  };

  const ensureGrid = (s: SheetMeta, rows: number, cols: number) => {
    if (rows > s.rowCount || cols > s.columnCount) {
      s.rowCount = Math.max(s.rowCount, rows);
      s.columnCount = Math.max(s.columnCount, cols);
      requests.push({
        updateSheetProperties: {
          properties: {
            sheetId: s.sheetId,
            gridProperties: { rowCount: s.rowCount, columnCount: s.columnCount },
          },
          fields: "gridProperties.rowCount,gridProperties.columnCount",
        },
      });
    }
  };

  const countCells = (values: unknown[][]) => {
    totalCells += values.reduce((n, r) => n + r.length, 0);
    if (totalCells > MAX_TOTAL_CELLS) {
      throw new UserFacingError(`Trop de cellules dans une meme requete (max ${MAX_TOTAL_CELLS}). Decoupez en plusieurs appels.`);
    }
  };

  for (const op of ops) {
    switch (op.op) {
      case "write_values": {
        const s = sheet(op.sheetName);
        const start = parseA1Range(op.startCell);
        const width = Math.max(...op.values.map((r) => r.length));
        countCells(op.values);
        ensureGrid(s, start.startRow! - 1 + op.values.length, start.startCol! - 1 + width);
        requests.push({
          updateCells: {
            rows: toRows(op.values, op.formulas),
            fields: "userEnteredValue",
            start: { sheetId: s.sheetId, rowIndex: start.startRow! - 1, columnIndex: start.startCol! - 1 },
          },
        });
        plan.push(`ecrit ${op.values.length}x${width} cellules dans "${s.title}" a partir de ${op.startCell.toUpperCase()}`);
        break;
      }
      case "append_rows": {
        const s = sheet(op.sheetName);
        countCells(op.rows);
        requests.push({
          appendCells: { sheetId: s.sheetId, rows: toRows(op.rows, op.formulas), fields: "userEnteredValue" },
        });
        plan.push(`ajoute ${op.rows.length} ligne(s) a la fin de "${s.title}"`);
        break;
      }
      case "clear_range": {
        const s = sheet(op.sheetName);
        requests.push({ repeatCell: { range: gridRange(s.sheetId, op.range), fields: "userEnteredValue" } });
        plan.push(`vide les valeurs de ${op.range.toUpperCase()} dans "${s.title}"`);
        break;
      }
      case "format_range": {
        const s = sheet(op.sheetName);
        const fmt: sheets_v4.Schema$CellFormat = {};
        const fields: string[] = [];
        const text: sheets_v4.Schema$TextFormat = {};
        if (op.bold !== undefined) (text.bold = op.bold), fields.push("userEnteredFormat.textFormat.bold");
        if (op.italic !== undefined) (text.italic = op.italic), fields.push("userEnteredFormat.textFormat.italic");
        if (op.fontSize !== undefined) (text.fontSize = op.fontSize), fields.push("userEnteredFormat.textFormat.fontSize");
        if (op.textColor) (text.foregroundColor = rgb(op.textColor)), fields.push("userEnteredFormat.textFormat.foregroundColor");
        if (Object.keys(text).length) fmt.textFormat = text;
        if (op.backgroundColor) (fmt.backgroundColor = rgb(op.backgroundColor)), fields.push("userEnteredFormat.backgroundColor");
        if (op.horizontalAlignment) (fmt.horizontalAlignment = op.horizontalAlignment), fields.push("userEnteredFormat.horizontalAlignment");
        if (op.wrapText !== undefined) (fmt.wrapStrategy = op.wrapText ? "WRAP" : "OVERFLOW_CELL"), fields.push("userEnteredFormat.wrapStrategy");
        if (op.numberFormat) (fmt.numberFormat = op.numberFormat), fields.push("userEnteredFormat.numberFormat");
        if (!fields.length) throw new UserFacingError("format_range : aucune propriete de mise en forme fournie.");
        requests.push({
          repeatCell: {
            range: gridRange(s.sheetId, op.range),
            cell: { userEnteredFormat: fmt },
            fields: fields.join(","),
          },
        });
        plan.push(`met en forme ${op.range.toUpperCase()} dans "${s.title}"`);
        break;
      }
      case "add_sheet": {
        if (model.has(op.sheetName)) throw new UserFacingError(`Un onglet "${op.sheetName}" existe deja.`);
        const meta: SheetMeta = { sheetId: nextId++, title: op.sheetName, rowCount: op.rows, columnCount: op.columns };
        model.set(meta.title, meta);
        requests.push({
          addSheet: {
            properties: {
              sheetId: meta.sheetId,
              title: meta.title,
              gridProperties: { rowCount: meta.rowCount, columnCount: meta.columnCount },
            },
          },
        });
        plan.push(`ajoute l'onglet "${op.sheetName}"`);
        break;
      }
      case "rename_sheet": {
        const s = sheet(op.sheetName);
        if (model.has(op.newName) && op.newName !== op.sheetName) throw new UserFacingError(`Un onglet "${op.newName}" existe deja.`);
        model.delete(s.title);
        s.title = op.newName;
        model.set(s.title, s);
        requests.push({ updateSheetProperties: { properties: { sheetId: s.sheetId, title: op.newName }, fields: "title" } });
        plan.push(`renomme l'onglet "${op.sheetName}" en "${op.newName}"`);
        break;
      }
      case "delete_sheet": {
        const s = sheet(op.sheetName);
        if (model.size <= 1) throw new UserFacingError("Impossible de supprimer le dernier onglet du classeur.");
        model.delete(s.title);
        requests.push({ deleteSheet: { sheetId: s.sheetId } });
        plan.push(`SUPPRIME l'onglet "${op.sheetName}"`);
        break;
      }
      case "insert_rows": {
        const s = sheet(op.sheetName);
        s.rowCount += op.count;
        requests.push({
          insertDimension: {
            range: { sheetId: s.sheetId, dimension: "ROWS", startIndex: op.beforeRow - 1, endIndex: op.beforeRow - 1 + op.count },
            inheritFromBefore: op.beforeRow > 1,
          },
        });
        plan.push(`insere ${op.count} ligne(s) avant la ligne ${op.beforeRow} de "${s.title}"`);
        break;
      }
      case "delete_rows": {
        const s = sheet(op.sheetName);
        s.rowCount = Math.max(0, s.rowCount - op.count);
        requests.push({
          deleteDimension: {
            range: { sheetId: s.sheetId, dimension: "ROWS", startIndex: op.startRow - 1, endIndex: op.startRow - 1 + op.count },
          },
        });
        plan.push(`supprime ${op.count} ligne(s) a partir de la ligne ${op.startRow} de "${s.title}"`);
        break;
      }
      case "insert_columns": {
        const s = sheet(op.sheetName);
        const at = columnNumber(op.beforeColumn);
        s.columnCount += op.count;
        requests.push({
          insertDimension: {
            range: { sheetId: s.sheetId, dimension: "COLUMNS", startIndex: at - 1, endIndex: at - 1 + op.count },
            inheritFromBefore: at > 1,
          },
        });
        plan.push(`insere ${op.count} colonne(s) avant ${columnLetters(at)} dans "${s.title}"`);
        break;
      }
      case "delete_columns": {
        const s = sheet(op.sheetName);
        const at = columnNumber(op.startColumn);
        s.columnCount = Math.max(0, s.columnCount - op.count);
        requests.push({
          deleteDimension: {
            range: { sheetId: s.sheetId, dimension: "COLUMNS", startIndex: at - 1, endIndex: at - 1 + op.count },
          },
        });
        plan.push(`supprime ${op.count} colonne(s) a partir de ${columnLetters(at)} dans "${s.title}"`);
        break;
      }
    }
  }
  return { requests, plan };
}
