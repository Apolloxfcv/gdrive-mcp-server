import { posix } from "node:path";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import type JSZip from "jszip";
import { loadZipSafely } from "./office-utils";
import { UserFacingError } from "./tool-guards";

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
