import { posix } from "node:path";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import type JSZip from "jszip";
import { loadZipSafely } from "./office-utils";
import { UserFacingError } from "./tool-guards";
import { columnLetters, parseA1Range } from "./a1";

/**
 * Modification de cellules d'un .xlsx par patch XML cible.
 *
 * Un aller-retour ExcelJS (load + writeBuffer) relit puis reecrit TOUT le
 * classeur et perd ou abime ce qu'il ne connait pas (graphiques, images,
 * tableaux croises, mises en forme conditionnelles, plages nommees...). Ici,
 * seul le XML de la feuille visee est modifie, et uniquement les cellules
 * demandees : toutes les autres parties du ZIP restent identiques.
 *
 * Le fichier est toujours re-uploade en entier (l'API Drive n'a pas d'ecriture
 * partielle), mais son contenu ne differe que des cellules modifiees.
 */

export type XlsxCellUpdate = { cell: string; value: string | number | boolean };

const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_XML = "http://www.w3.org/XML/1998/namespace";

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, "text/xml") as unknown as Document;
}

function serializeXml(doc: Document): string {
  const xml = new XMLSerializer().serializeToString(doc as any);
  return xml.startsWith("<?xml")
    ? xml
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`;
}

function childElements(node: Node, localName: string): Element[] {
  const out: Element[] = [];
  for (let n = node.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && (n as Element).localName === localName) out.push(n as Element);
  }
  return out;
}

function columnNumber(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function splitCellRef(ref: string): { col: number; row: number } {
  const m = /^([A-Z]+)([0-9]+)$/i.exec(ref);
  if (!m) throw new UserFacingError(`Reference de cellule invalide: ${ref}`);
  return { col: columnNumber(m[1]), row: Number(m[2]) };
}

async function readZipXml(zip: JSZip, path: string): Promise<Document | undefined> {
  const xml = await zip.file(path)?.async("string");
  return xml === undefined ? undefined : parseXml(xml);
}

/** Retrouve le chemin du XML d'une feuille a partir de son nom (workbook.xml + rels). */
async function locateSheet(zip: JSZip, sheetName: string) {
  const workbook = await readZipXml(zip, "xl/workbook.xml");
  const rels = await readZipXml(zip, "xl/_rels/workbook.xml.rels");
  if (!workbook || !rels) throw new UserFacingError("Structure .xlsx invalide (workbook.xml introuvable).");

  const sheets = Array.from(workbook.getElementsByTagNameNS(NS_MAIN, "sheet"));
  const sheet = sheets.find((s) => s.getAttribute("name") === sheetName);
  if (!sheet) {
    throw new UserFacingError(
      `Feuille "${sheetName}" introuvable. Feuilles disponibles: ${sheets
        .map((s) => s.getAttribute("name"))
        .join(", ")}`
    );
  }

  const rid = sheet.getAttributeNS(NS_REL, "id");
  const rel = Array.from(rels.getElementsByTagNameNS(NS_PKG_REL, "Relationship")).find(
    (r) => r.getAttribute("Id") === rid
  );
  const target = rel?.getAttribute("Target");
  if (!target) throw new UserFacingError(`Feuille "${sheetName}" illisible (relation manquante).`);

  const path = target.startsWith("/") ? target.slice(1) : posix.normalize(`xl/${target}`);
  return { workbook, rels, path };
}

/** Cellule existante ou nouvelle, inseree a sa place (lignes et colonnes triees). */
function getOrCreateCell(doc: Document, sheetData: Element, ref: string): Element {
  const { col, row } = splitCellRef(ref);
  const wanted = ref.toUpperCase();

  let rowEl: Element | undefined;
  let rowBefore: Element | undefined;
  for (const r of childElements(sheetData, "row")) {
    const n = Number(r.getAttribute("r"));
    if (n === row) {
      rowEl = r;
      break;
    }
    if (n > row) {
      rowBefore = r;
      break;
    }
  }
  if (!rowEl) {
    rowEl = doc.createElementNS(NS_MAIN, "row");
    rowEl.setAttribute("r", String(row));
    sheetData.insertBefore(rowEl, rowBefore ?? null);
  }
  // `spans` n'est qu'un indice d'optimisation ; on le retire plutot que de le laisser faux.
  rowEl.removeAttribute("spans");

  let cellBefore: Element | undefined;
  for (const c of childElements(rowEl, "c")) {
    const cRef = (c.getAttribute("r") ?? "").toUpperCase();
    if (cRef === wanted) return c;
    if (columnNumber(cRef.replace(/[0-9]/g, "")) > col) {
      cellBefore = c;
      break;
    }
  }
  const cell = doc.createElementNS(NS_MAIN, "c");
  cell.setAttribute("r", wanted);
  rowEl.insertBefore(cell, cellBefore ?? null);
  return cell;
}

/** Ecrit une valeur en gardant le style (`s`) ; retourne true si une formule a ete ecrasee. */
function writeCellValue(doc: Document, cell: Element, value: XlsxCellUpdate["value"]): boolean {
  const formula = childElements(cell, "f")[0];
  if (formula) {
    const kind = formula.getAttribute("t");
    // Ecraser le maitre d'une formule partagee/matricielle corromprait les cellules dependantes.
    if (kind === "array" || (kind === "shared" && formula.hasAttribute("ref"))) {
      throw new UserFacingError(
        `La cellule ${cell.getAttribute("r")} contient une formule partagee ou matricielle : modification refusee pour ne pas corrompre le fichier.`
      );
    }
  }

  for (const name of Array.from(cell.attributes).map((a) => a.name)) {
    if (name !== "r" && name !== "s") cell.removeAttribute(name);
  }
  while (cell.firstChild) cell.removeChild(cell.firstChild);

  const text = (tag: string, content: string) => {
    const el = doc.createElementNS(NS_MAIN, tag);
    el.appendChild(doc.createTextNode(content));
    return el;
  };

  if (typeof value === "number") {
    if (!Number.isFinite(value)) throw new UserFacingError("Valeur numerique invalide (NaN/Infinity).");
    cell.appendChild(text("v", String(value)));
  } else if (typeof value === "boolean") {
    cell.setAttribute("t", "b");
    cell.appendChild(text("v", value ? "1" : "0"));
  } else {
    // Chaine inline : evite de toucher a sharedStrings.xml (Excel la normalisera a la prochaine sauvegarde).
    cell.setAttribute("t", "inlineStr");
    const is = doc.createElementNS(NS_MAIN, "is");
    const t = text("t", value.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, ""));
    t.setAttributeNS(NS_XML, "xml:space", "preserve");
    is.appendChild(t);
    cell.appendChild(is);
  }
  return Boolean(formula);
}

export async function updateXlsxCells(
  buffer: Buffer,
  sheetName: string,
  updates: XlsxCellUpdate[]
): Promise<Buffer> {
  const zip = await loadZipSafely(buffer);
  const { workbook, rels, path: sheetPath } = await locateSheet(zip, sheetName);

  const sheetXml = await readZipXml(zip, sheetPath);
  if (!sheetXml) throw new UserFacingError(`Feuille "${sheetName}" introuvable dans l'archive.`);
  const sheetData = sheetXml.getElementsByTagNameNS(NS_MAIN, "sheetData")[0];
  if (!sheetData) throw new UserFacingError(`Feuille "${sheetName}" invalide (sheetData manquant).`);

  let formulaOverwritten = false;
  for (const { cell, value } of updates) {
    const el = getOrCreateCell(sheetXml, sheetData, cell);
    formulaOverwritten = writeCellValue(sheetXml, el, value) || formulaOverwritten;
  }
  zip.file(sheetPath, serializeXml(sheetXml));

  // Les formules dependantes doivent etre recalculees a l'ouverture (valeurs en cache perimees).
  const calcPr = workbook.getElementsByTagNameNS(NS_MAIN, "calcPr")[0];
  if (calcPr) {
    calcPr.setAttribute("fullCalcOnLoad", "1");
    zip.file("xl/workbook.xml", serializeXml(workbook));
  }

  // calcChain.xml reference les cellules a formule : s'il pointe vers une formule
  // ecrasee, Excel signale un fichier corrompu. On le supprime (Excel le reconstruit).
  if (formulaOverwritten && zip.file("xl/calcChain.xml")) {
    zip.remove("xl/calcChain.xml");

    for (const r of Array.from(rels.getElementsByTagNameNS(NS_PKG_REL, "Relationship"))) {
      if ((r.getAttribute("Type") ?? "").endsWith("/calcChain")) r.parentNode?.removeChild(r);
    }
    zip.file("xl/_rels/workbook.xml.rels", serializeXml(rels));

    const types = await readZipXml(zip, "[Content_Types].xml");
    if (types) {
      for (const o of Array.from(types.getElementsByTagNameNS(NS_CONTENT_TYPES, "Override"))) {
        if (o.getAttribute("PartName") === "/xl/calcChain.xml") o.parentNode?.removeChild(o);
      }
      zip.file("[Content_Types].xml", serializeXml(types));
    }
  }

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

// ---------- Edition par lot (plusieurs feuilles, formules, ajout de lignes, effacement) ----------

export type XlsxBatchOp =
  | { op: "set"; sheetName: string; cell: string; value: string | number | boolean | null }
  | { op: "formula"; sheetName: string; cell: string; formula: string }
  | { op: "clear_range"; sheetName: string; range: string }
  | { op: "append_rows"; sheetName: string; rows: (string | number | boolean | null)[][]; startColumn?: string };

const WORKBOOK_BEFORE_CALCPR = ["sheets", "functionGroups", "externalReferences", "definedNames"];

function assertWritable(cell: Element) {
  const formula = childElements(cell, "f")[0];
  if (!formula) return;
  const kind = formula.getAttribute("t");
  if (kind === "array" || (kind === "shared" && formula.hasAttribute("ref"))) {
    throw new UserFacingError(
      `La cellule ${cell.getAttribute("r")} contient une formule partagee ou matricielle : modification refusee pour ne pas corrompre le fichier.`
    );
  }
}

function cellHasContent(cell: Element): boolean {
  return (
    childElements(cell, "v").length > 0 ||
    childElements(cell, "is").length > 0 ||
    childElements(cell, "f").length > 0
  );
}

/** Vide la cellule en gardant son style ; retourne true si elle contenait une formule. */
function clearCell(cell: Element): boolean {
  assertWritable(cell);
  const hadFormula = childElements(cell, "f").length > 0;
  for (const name of Array.from(cell.attributes).map((a) => a.name)) {
    if (name !== "r" && name !== "s") cell.removeAttribute(name);
  }
  while (cell.firstChild) cell.removeChild(cell.firstChild);
  return hadFormula;
}

function writeFormula(doc: Document, cell: Element, formula: string) {
  clearCell(cell);
  const f = doc.createElementNS(NS_MAIN, "f");
  const clean = formula.replace(/^=/, "").replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  f.appendChild(doc.createTextNode(clean));
  cell.appendChild(f);
}

/** Refuse d'ecrire dans une cellule couverte par une fusion (hors cellule en haut a gauche). */
function assertNotInsideMerge(sheetXml: Document, ref: string) {
  const { col, row } = splitCellRef(ref);
  for (const m of Array.from(sheetXml.getElementsByTagNameNS(NS_MAIN, "mergeCell"))) {
    const [a, b] = (m.getAttribute("ref") ?? "").split(":");
    if (!a || !b) continue;
    const from = splitCellRef(a);
    const to = splitCellRef(b);
    const inside = col >= from.col && col <= to.col && row >= from.row && row <= to.row;
    if (inside && !(col === from.col && row === from.row)) {
      throw new UserFacingError(
        `La cellule ${ref.toUpperCase()} fait partie de la fusion ${m.getAttribute("ref")} : ecrire dans ${a} (cellule en haut a gauche).`
      );
    }
  }
}

function lastUsedRow(sheetData: Element): number {
  let last = 0;
  for (const r of childElements(sheetData, "row")) {
    if (childElements(r, "c").some(cellHasContent)) last = Math.max(last, Number(r.getAttribute("r")));
  }
  return last;
}

function ensureFullCalcOnLoad(workbook: Document) {
  let calcPr = workbook.getElementsByTagNameNS(NS_MAIN, "calcPr")[0];
  if (!calcPr) {
    const root = workbook.documentElement;
    calcPr = workbook.createElementNS(NS_MAIN, "calcPr");
    // Ordre impose par le schema : calcPr vient apres sheets/functionGroups/externalReferences/definedNames.
    let anchor: Element | undefined;
    for (const name of WORKBOOK_BEFORE_CALCPR) anchor = childElements(root, name)[0] ?? anchor;
    root.insertBefore(calcPr, anchor ? anchor.nextSibling : null);
  }
  calcPr.setAttribute("fullCalcOnLoad", "1");
}

/** Met a jour <dimension> (simple indice, mais certains lecteurs s'y fient) apres ajout de cellules. */
function refreshDimension(sheetXml: Document, sheetData: Element) {
  const dim = sheetXml.getElementsByTagNameNS(NS_MAIN, "dimension")[0];
  if (!dim) return;
  let minC = Infinity, maxC = 0, minR = Infinity, maxR = 0;
  for (const r of childElements(sheetData, "row")) {
    for (const c of childElements(r, "c")) {
      const { col, row } = splitCellRef(c.getAttribute("r") ?? "A1");
      minC = Math.min(minC, col);
      maxC = Math.max(maxC, col);
      minR = Math.min(minR, row);
      maxR = Math.max(maxR, row);
    }
  }
  if (maxC === 0) return;
  const from = `${columnLetters(minC)}${minR}`;
  const to = `${columnLetters(maxC)}${maxR}`;
  dim.setAttribute("ref", from === to ? from : `${from}:${to}`);
}

/**
 * Applique une liste d'operations sur un ou plusieurs onglets d'un .xlsx en un
 * seul passage (un telechargement, un upload). Seules les cellules visees sont
 * touchees ; tout le reste de l'archive est recopie tel quel.
 */
export async function editXlsxWorkbook(
  buffer: Buffer,
  ops: XlsxBatchOp[]
): Promise<{ buffer: Buffer; warnings: string[]; cellsChanged: number }> {
  const zip = await loadZipSafely(buffer);
  const sheets = new Map<string, { path: string; xml: Document; sheetData: Element }>();
  const warnings: string[] = [];
  let formulaTouched = false;
  let cellsChanged = 0;
  let workbookDoc: Document | undefined;
  let relsDoc: Document | undefined;

  const open = async (sheetName: string) => {
    const cached = sheets.get(sheetName);
    if (cached) return cached;
    const { workbook, rels, path } = await locateSheet(zip, sheetName);
    workbookDoc = workbook;
    relsDoc = rels;
    const xml = await readZipXml(zip, path);
    if (!xml) throw new UserFacingError(`Feuille "${sheetName}" introuvable dans l'archive.`);
    const sheetData = xml.getElementsByTagNameNS(NS_MAIN, "sheetData")[0];
    if (!sheetData) throw new UserFacingError(`Feuille "${sheetName}" invalide (sheetData manquant).`);
    if (xml.getElementsByTagNameNS(NS_MAIN, "tableParts").length) {
      warnings.push(
        `"${sheetName}" contient un tableau Excel formate : ne modifiez pas ses cellules d'en-tete (Excel signalerait un fichier a reparer).`
      );
    }
    const entry = { path, xml, sheetData };
    sheets.set(sheetName, entry);
    return entry;
  };

  const setCell = (
    s: { xml: Document; sheetData: Element },
    ref: string,
    value: string | number | boolean | null
  ) => {
    assertNotInsideMerge(s.xml, ref);
    const el = getOrCreateCell(s.xml, s.sheetData, ref);
    formulaTouched = (value === null ? clearCell(el) : writeCellValue(s.xml, el, value)) || formulaTouched;
    cellsChanged++;
  };

  for (const op of ops) {
    const s = await open(op.sheetName);
    if (op.op === "set") {
      setCell(s, op.cell, op.value);
    } else if (op.op === "formula") {
      assertNotInsideMerge(s.xml, op.cell);
      writeFormula(s.xml, getOrCreateCell(s.xml, s.sheetData, op.cell), op.formula);
      formulaTouched = true;
      cellsChanged++;
    } else if (op.op === "clear_range") {
      const range = parseA1Range(op.range);
      for (const r of childElements(s.sheetData, "row")) {
        const n = Number(r.getAttribute("r"));
        if ((range.startRow !== undefined && n < range.startRow) || (range.endRow !== undefined && n > range.endRow)) continue;
        for (const c of childElements(r, "c")) {
          const col = columnNumber((c.getAttribute("r") ?? "").replace(/[0-9]/g, ""));
          if ((range.startCol !== undefined && col < range.startCol) || (range.endCol !== undefined && col > range.endCol)) continue;
          if (!cellHasContent(c)) continue;
          formulaTouched = clearCell(c) || formulaTouched;
          cellsChanged++;
        }
      }
    } else {
      let row = lastUsedRow(s.sheetData) + 1;
      const startCol = columnNumber(op.startColumn ?? "A");
      for (const values of op.rows) {
        values.forEach((value, i) => {
          if (value !== null) setCell(s, `${columnLetters(startCol + i)}${row}`, value);
        });
        row++;
      }
    }
  }

  for (const { path, xml, sheetData } of sheets.values()) {
    refreshDimension(xml, sheetData);
    zip.file(path, serializeXml(xml));
  }

  // Valeurs en cache perimees : recalcul complet a l'ouverture.
  if (workbookDoc) {
    const existingCalcPr = workbookDoc.getElementsByTagNameNS(NS_MAIN, "calcPr")[0];
    if (formulaTouched) ensureFullCalcOnLoad(workbookDoc);
    else existingCalcPr?.setAttribute("fullCalcOnLoad", "1");
    if (formulaTouched || existingCalcPr) zip.file("xl/workbook.xml", serializeXml(workbookDoc));
  }

  // Une formule ajoutee/retiree rend calcChain.xml inexact : Excel le reconstruit s'il est absent.
  if (formulaTouched && zip.file("xl/calcChain.xml") && relsDoc) {
    zip.remove("xl/calcChain.xml");
    for (const r of Array.from(relsDoc.getElementsByTagNameNS(NS_PKG_REL, "Relationship"))) {
      if ((r.getAttribute("Type") ?? "").endsWith("/calcChain")) r.parentNode?.removeChild(r);
    }
    zip.file("xl/_rels/workbook.xml.rels", serializeXml(relsDoc));
    const types = await readZipXml(zip, "[Content_Types].xml");
    if (types) {
      for (const o of Array.from(types.getElementsByTagNameNS(NS_CONTENT_TYPES, "Override"))) {
        if (o.getAttribute("PartName") === "/xl/calcChain.xml") o.parentNode?.removeChild(o);
      }
      zip.file("[Content_Types].xml", serializeXml(types));
    }
  }

  const out = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  return { buffer: out, warnings, cellsChanged };
}
