import { UserFacingError } from "./tool-guards";

/**
 * Notation A1 : conversions colonnes <-> numeros et analyse de plages
 * (cellule, plage, colonnes entieres "A:C", lignes entieres "2:10").
 * Tous les numeros de ligne/colonne sont 1-based ; les bornes absentes
 * (colonnes/lignes entieres) sont `undefined`.
 */

export function columnNumber(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function columnLetters(n: number): string {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export type A1Range = {
  startCol?: number;
  endCol?: number;
  startRow?: number;
  endRow?: number;
};

const CELL = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/i;
const COLS = /^([A-Z]{1,3}):([A-Z]{1,3})$/i;
const ROWS = /^([1-9][0-9]{0,6}):([1-9][0-9]{0,6})$/;

/** Plage A1 sans nom de feuille : "B3", "B3:D9", "A:C", "2:10", "B3:D". */
export const A1_RANGE_PATTERN =
  /^([A-Z]{1,3}[1-9][0-9]{0,6}(:[A-Z]{1,3}([1-9][0-9]{0,6})?)?|[A-Z]{1,3}:[A-Z]{1,3}|[1-9][0-9]{0,6}:[1-9][0-9]{0,6})$/i;

export function parseA1Range(input: string): A1Range {
  const ref = input.trim();
  const bad = () => new UserFacingError(`Plage A1 invalide: ${input}`);

  const single = CELL.exec(ref);
  if (single) {
    const col = columnNumber(single[1]);
    const row = Number(single[2]);
    return { startCol: col, endCol: col, startRow: row, endRow: row };
  }
  const cols = COLS.exec(ref);
  if (cols) {
    const a = columnNumber(cols[1]);
    const b = columnNumber(cols[2]);
    return { startCol: Math.min(a, b), endCol: Math.max(a, b) };
  }
  const rows = ROWS.exec(ref);
  if (rows) {
    const a = Number(rows[1]);
    const b = Number(rows[2]);
    return { startRow: Math.min(a, b), endRow: Math.max(a, b) };
  }
  const [from, to] = ref.split(":");
  const f = CELL.exec(from ?? "");
  if (f && to) {
    const t = CELL.exec(to) ?? /^([A-Z]{1,3})$/i.exec(to);
    if (!t) throw bad();
    const c1 = columnNumber(f[1]);
    const c2 = columnNumber(t[1]);
    const r1 = Number(f[2]);
    const r2 = t[2] ? Number(t[2]) : undefined;
    return {
      startCol: Math.min(c1, c2),
      endCol: Math.max(c1, c2),
      startRow: r2 === undefined ? r1 : Math.min(r1, r2),
      endRow: r2 === undefined ? undefined : Math.max(r1, r2),
    };
  }
  throw bad();
}
