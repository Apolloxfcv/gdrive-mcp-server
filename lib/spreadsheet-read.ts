import ExcelJS from "exceljs";
import { columnLetters, parseA1Range } from "./a1";
import { loadZipSafely } from "./office-utils";
import { UserFacingError } from "./tool-guards";

/**
 * Lecture "propre" d'un .xlsx : les objets ExcelJS (formules, texte riche,
 * liens, dates, erreurs) sont ramenes a des valeurs simples, et chaque feuille
 * est renvoyee alignee sur les coordonnees A1 (`startCell` + `rows`), de facon
 * a pouvoir viser directement les memes references dans les outils d'edition.
 */

export type CellValue = string | number | boolean | null;

export type SheetGrid = {
  name: string;
  hidden: boolean;
  /** Derniere ligne / colonne contenant des donnees dans la feuille entiere. */
  rowCount: number;
  columnCount: number;
  /** Absent en mode metadata seule. */
  startCell?: string;
  rows?: CellValue[][];
  truncated?: boolean;
};

export type ReadOptions = {
  sheetName?: string;
  range?: string;
  includeFormulas: boolean;
  maxRows: number;
  metadataOnly: boolean;
};

function normalize(value: unknown, includeFormula: boolean, formula?: string): CellValue {
  if (includeFormula && formula) return `=${formula}`;
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return value.toISOString();
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
  if (typeof value === "object") {
    const v = value as Record<string, unknown>;
    if ("result" in v) return normalize(v.result, false);
    if ("richText" in v && Array.isArray(v.richText)) {
      return (v.richText as { text?: string }[]).map((r) => r.text ?? "").join("");
    }
    if ("error" in v) return String(v.error);
    if ("text" in v) return normalize(v.text, false);
    if ("formula" in v || "sharedFormula" in v) return null; // formule sans valeur en cache
  }
  return String(value);
}

export async function readXlsxGrids(buffer: Buffer, opts: ReadOptions): Promise<SheetGrid[]> {
  await loadZipSafely(buffer);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer as any);

  if (opts.sheetName && !workbook.getWorksheet(opts.sheetName)) {
    throw new UserFacingError(
      `Feuille "${opts.sheetName}" introuvable. Feuilles disponibles: ${workbook.worksheets
        .map((w) => w.name)
        .join(", ")}`
    );
  }
  const range = opts.range ? parseA1Range(opts.range) : undefined;

  const out: SheetGrid[] = [];
  for (const ws of workbook.worksheets) {
    if (opts.sheetName && ws.name !== opts.sheetName) continue;
    const grid: SheetGrid = {
      name: ws.name,
      hidden: ws.state !== "visible",
      rowCount: ws.rowCount,
      columnCount: ws.columnCount,
    };
    if (!opts.metadataOnly) {
      const startRow = range?.startRow ?? 1;
      const wantedEnd = range?.endRow ?? ws.rowCount;
      const endRow = Math.min(wantedEnd, startRow + opts.maxRows - 1);
      const startCol = range?.startCol ?? 1;
      const endCol = Math.min(range?.endCol ?? ws.columnCount, startCol + 199);

      const rows: CellValue[][] = [];
      for (let r = startRow; r <= endRow; r++) {
        const row: CellValue[] = [];
        const wsRow = ws.getRow(r);
        for (let c = startCol; c <= endCol; c++) {
          const cell = wsRow.getCell(c);
          const formula = (cell as unknown as { formula?: string }).formula;
          row.push(normalize(cell.value, opts.includeFormulas, formula));
        }
        rows.push(row);
      }
      grid.startCell = `${columnLetters(startCol)}${startRow}`;
      grid.rows = rows;
      grid.truncated = endRow < wantedEnd;
    }
    out.push(grid);
  }
  return out;
}
