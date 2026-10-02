import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { z } from "zod";
import { loadZipSafely } from "./office-utils";
import { UserFacingError } from "./tool-guards";

/**
 * Lecture indexee et edition ciblee d'un .docx par patch XML de word/document.xml.
 *
 * Contrairement a une regeneration complete du document (Packer/Document), seuls
 * les paragraphes vises sont modifies : styles, en-tetes/pieds de page, images,
 * tableaux, numerotation, sections et toutes les autres parties du ZIP restent
 * identiques octet pour octet.
 *
 * Numerotation : l'index d'un paragraphe est sa position (0-based) parmi tous les
 * <w:p> du corps, tableaux compris, dans l'ordre du document ; les contenus de
 * secours mc:Fallback sont ignores. Lecture et edition utilisent le meme
 * parcours (`listParagraphs`), et les `index` d'une liste d'operations se
 * rapportent toujours au document tel qu'il a ete lu (avant toute operation).
 */

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_XML = "http://www.w3.org/XML/1998/namespace";
const MAX_OPS = 500;

// ---------- Schema des operations ----------

const target = z
  .object({
    index: z.number().int().min(0).optional().describe("Index du paragraphe (voir drive_read_docx_paragraphs)"),
    containsText: z.string().min(1).max(500).optional().describe("Texte contenu par UN SEUL paragraphe"),
  })
  .refine((t) => (t.index === undefined) !== (t.containsText === undefined), {
    message: "Fournir exactement un de : index, containsText",
  });

const oneLine = z
  .string()
  .max(100_000)
  .refine((s) => !/[\r\n]/.test(s), { message: "Pas de saut de ligne : un paragraphe par element" });

export const docxOpSchema = z.discriminatedUnion("op", [
  z.object({
    op: z.literal("replace_text"),
    find: z.string().min(1).max(2000),
    replace: oneLine,
    occurrence: z.enum(["all", "first"]).default("all"),
    required: z.boolean().default(true).describe("true : erreur (rien n'est ecrit) si le texte est introuvable"),
  }),
  z.object({ op: z.literal("replace_paragraph"), target, text: oneLine }),
  z.object({
    op: z.literal("insert_paragraphs"),
    target,
    position: z.enum(["before", "after"]).default("after"),
    paragraphs: z.array(oneLine).min(1).max(200),
    copyFormatting: z.boolean().default(true).describe("Reprend le style (titre, puce...) du paragraphe cible"),
  }),
  z.object({
    op: z.literal("append_paragraphs"),
    paragraphs: z.array(oneLine).min(1).max(200),
  }),
  z.object({ op: z.literal("delete_paragraph"), target }),
]);

export type DocxOp = z.infer<typeof docxOpSchema>;
export const docxOpsSchema = z.array(docxOpSchema).min(1).max(MAX_OPS);

// ---------- XML ----------

function parse(xml: string): Document {
  return new DOMParser().parseFromString(xml, "text/xml") as unknown as Document;
}

function serialize(doc: Document): string {
  const xml = new XMLSerializer().serializeToString(doc as any);
  return xml.startsWith("<?xml") ? xml : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`;
}

async function loadDocument(buffer: Buffer) {
  const zip = await loadZipSafely(buffer);
  const xml = await zip.file("word/document.xml")?.async("string");
  if (!xml) throw new UserFacingError("document.xml introuvable dans ce .docx");
  return { zip, doc: parse(xml) };
}

function kids(node: Node, localName?: string): Element[] {
  const out: Element[] = [];
  for (let n = node.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && (!localName || (n as Element).localName === localName)) out.push(n as Element);
  }
  return out;
}

function inFallback(el: Element): boolean {
  for (let n = el.parentNode; n; n = n.parentNode) {
    if (n.nodeType === 1 && (n as Element).localName === "Fallback") return true;
  }
  return false;
}

function listParagraphs(doc: Document): Element[] {
  return Array.from(doc.getElementsByTagNameNS(W, "p")).filter((p) => !inFallback(p));
}

type Piece = { ch: string; t?: Element; off?: number };

/** Caracteres du paragraphe : ceux des <w:t> sont rattaches a leur noeud ; tab/saut de ligne ne sont pas editables. */
function pieces(p: Element): Piece[] {
  const out: Piece[] = [];
  const walk = (node: Element) => {
    for (const c of kids(node)) {
      switch (c.localName) {
        case "pPr":
        case "txbxContent": // paragraphes de zones de texte : listes a part
        case "Fallback":
          break;
        case "t": {
          const text = c.textContent ?? "";
          for (let i = 0; i < text.length; i++) out.push({ ch: text[i], t: c, off: i });
          break;
        }
        case "tab":
          out.push({ ch: "\t" });
          break;
        case "br":
        case "cr":
          out.push({ ch: "\n" });
          break;
        default:
          walk(c);
      }
    }
  };
  walk(p);
  return out;
}

function paragraphText(p: Element): string {
  return pieces(p)
    .map((x) => x.ch)
    .join("");
}

function paragraphStyle(p: Element): string | undefined {
  const pPr = kids(p, "pPr")[0];
  const style = pPr && kids(pPr, "pStyle")[0];
  return style?.getAttributeNS(W, "val") || style?.getAttribute("w:val") || undefined;
}

function inTable(p: Element): boolean {
  for (let n = p.parentNode; n; n = n.parentNode) {
    if (n.nodeType === 1 && (n as Element).localName === "tc") return true;
  }
  return false;
}

// ---------- Lecture ----------

export type DocxParagraph = { index: number; text: string; style?: string; inTable: boolean };

export async function readDocxParagraphs(buffer: Buffer): Promise<DocxParagraph[]> {
  const { doc } = await loadDocument(buffer);
  return listParagraphs(doc).map((p, index) => ({
    index,
    text: paragraphText(p),
    ...(paragraphStyle(p) ? { style: paragraphStyle(p) } : {}),
    inTable: inTable(p),
  }));
}

// ---------- Edition ----------

const CONTROL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g;

function setText(t: Element, text: string) {
  while (t.firstChild) t.removeChild(t.firstChild);
  t.appendChild(t.ownerDocument.createTextNode(text));
  t.setAttributeNS(NS_XML, "xml:space", "preserve");
}

/** Remplace `find` par `replace` dans un paragraphe, y compris quand le texte est coupe entre plusieurs runs. */
function replaceInParagraph(p: Element, find: string, replace: string, limit: number): number {
  const ps = pieces(p);
  const text = ps.map((x) => x.ch).join("");
  const spans: [number, number][] = [];
  for (let from = 0; spans.length < limit; ) {
    const at = text.indexOf(find, from);
    if (at === -1) break;
    // Un match qui recouvre un tab/saut de ligne n'est pas remplacable proprement.
    if (ps.slice(at, at + find.length).every((x) => x.t)) spans.push([at, at + find.length]);
    from = at + find.length;
  }

  // De la fin vers le debut : les decalages des matchs precedents restent valides.
  for (const [s, e] of spans.reverse()) {
    const first = ps[s];
    const last = ps[e - 1];
    const touched: Element[] = [];
    for (let i = s; i < e; i++) if (!touched.includes(ps[i].t!)) touched.push(ps[i].t!);

    for (const t of touched) {
      const cur = t.textContent ?? "";
      const from = t === first.t ? first.off! : 0;
      const to = t === last.t ? last.off! + 1 : cur.length;
      const insert = t === first.t ? replace : "";
      setText(t, cur.slice(0, from) + insert + cur.slice(to));
    }
  }
  return spans.length;
}

const DROP_IN_CLONE = new Set(["sectPr", "ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange"]);

function cloneClean(node: Element): Element {
  const copy = node.cloneNode(true) as Element;
  const strip = (el: Element) => {
    for (const c of kids(el)) {
      if (DROP_IN_CLONE.has(c.localName)) el.removeChild(c);
      else strip(c);
    }
  };
  strip(copy);
  return copy;
}

function makeParagraph(doc: Document, text: string, model?: Element): Element {
  const p = doc.createElementNS(W, "w:p");
  const clean = text.replace(CONTROL, "");

  const pPr = model && kids(model, "pPr")[0];
  if (pPr) p.appendChild(cloneClean(pPr));

  const r = doc.createElementNS(W, "w:r");
  const firstRun = model && Array.from(model.getElementsByTagNameNS(W, "r")).find((x) => kids(x, "rPr").length);
  if (firstRun) r.appendChild(cloneClean(kids(firstRun, "rPr")[0]));
  const t = doc.createElementNS(W, "w:t");
  setText(t, clean);
  r.appendChild(t);
  p.appendChild(r);
  return p;
}

type Ctx = { doc: Document; body: Element; original: Element[] };

function resolve(ctx: Ctx, t: { index?: number; containsText?: string }): Element {
  if (t.index !== undefined) {
    const p = ctx.original[t.index];
    if (!p) throw new UserFacingError(`Paragraphe d'index ${t.index} inexistant (le document en compte ${ctx.original.length}).`);
    if (!p.parentNode) throw new UserFacingError(`Le paragraphe d'index ${t.index} a deja ete supprime par une operation precedente.`);
    return p;
  }
  const needle = t.containsText!;
  const hits = listParagraphs(ctx.doc).filter((p) => paragraphText(p).includes(needle));
  if (hits.length === 0) throw new UserFacingError(`Aucun paragraphe ne contient "${needle}".`);
  if (hits.length > 1) {
    const idx = hits
      .slice(0, 8)
      .map((p) => ctx.original.indexOf(p))
      .map((i) => (i === -1 ? "(nouveau)" : String(i)));
    throw new UserFacingError(
      `"${needle}" est present dans ${hits.length} paragraphes (index ${idx.join(", ")}) : precisez avec "index".`
    );
  }
  return hits[0];
}

function assertSane(doc: Document, hadFinalSectPr: boolean) {
  const body = doc.getElementsByTagNameNS(W, "body")[0];
  if (hadFinalSectPr) {
    const last = kids(body).pop();
    if (last?.localName !== "sectPr") throw new UserFacingError("Edition annulee : la section finale du document serait invalide.");
  }
  for (const tc of Array.from(doc.getElementsByTagNameNS(W, "tc"))) {
    if (kids(tc).pop()?.localName !== "p") {
      throw new UserFacingError("Edition annulee : une cellule de tableau ne se terminerait plus par un paragraphe.");
    }
  }
}

export async function editDocx(
  buffer: Buffer,
  ops: DocxOp[]
): Promise<{ buffer: Buffer; report: string[] }> {
  const { zip, doc } = await loadDocument(buffer);
  const body = doc.getElementsByTagNameNS(W, "body")[0];
  if (!body) throw new UserFacingError("Structure .docx invalide (w:body manquant).");
  const hadFinalSectPr = kids(body).pop()?.localName === "sectPr";
  const ctx: Ctx = { doc, body, original: listParagraphs(doc) };
  const report: string[] = [];

  for (const op of ops) {
    switch (op.op) {
      case "replace_text": {
        const find = op.find.replace(CONTROL, "");
        const replace = op.replace.replace(CONTROL, "");
        let left = op.occurrence === "first" ? 1 : Infinity;
        let done = 0;
        for (const p of listParagraphs(doc)) {
          if (left <= 0) break;
          const n = replaceInParagraph(p, find, replace, left);
          done += n;
          left -= n;
        }
        if (done === 0 && op.required) {
          throw new UserFacingError(
            `Texte "${op.find.slice(0, 80)}" introuvable : rien n'a ete modifie. (Verifiez avec drive_read_docx_paragraphs ; la recherche est sensible a la casse.)`
          );
        }
        report.push(`replace_text "${op.find.slice(0, 40)}" : ${done} remplacement(s)`);
        break;
      }
      case "replace_paragraph": {
        const p = resolve(ctx, op.target);
        const ts = pieces(p)
          .map((x) => x.t)
          .filter((t, i, a): t is Element => !!t && a.indexOf(t) === i);
        const text = op.text.replace(CONTROL, "");
        if (ts.length) {
          setText(ts[0], text);
          for (const t of ts.slice(1)) setText(t, "");
        } else {
          const r = doc.createElementNS(W, "w:r");
          const t = doc.createElementNS(W, "w:t");
          setText(t, text);
          r.appendChild(t);
          p.appendChild(r);
        }
        report.push(`replace_paragraph : paragraphe modifie`);
        break;
      }
      case "insert_paragraphs": {
        const ref = resolve(ctx, op.target);
        const parent = ref.parentNode!;
        const model = op.copyFormatting ? ref : undefined;
        const anchor = op.position === "after" ? ref.nextSibling : ref;
        for (const text of op.paragraphs) parent.insertBefore(makeParagraph(doc, text, model), anchor);
        report.push(`insert_paragraphs : ${op.paragraphs.length} paragraphe(s) ${op.position === "after" ? "apres" : "avant"} la cible`);
        break;
      }
      case "append_paragraphs": {
        const last = kids(body).pop();
        const anchor = last?.localName === "sectPr" ? last : null; // jamais apres la section finale
        for (const text of op.paragraphs) body.insertBefore(makeParagraph(doc, text), anchor);
        report.push(`append_paragraphs : ${op.paragraphs.length} paragraphe(s) ajoute(s) a la fin`);
        break;
      }
      case "delete_paragraph": {
        const p = resolve(ctx, op.target);
        const pPr = kids(p, "pPr")[0];
        if (pPr && kids(pPr, "sectPr").length) {
          throw new UserFacingError("Ce paragraphe porte un saut de section : suppression refusee.");
        }
        const parent = p.parentNode as Element;
        if (parent.localName === "tc") {
          // Une cellule doit toujours se terminer par un paragraphe.
          const rest = kids(parent).filter((k) => k !== p);
          if (rest[rest.length - 1]?.localName !== "p") {
            throw new UserFacingError(
              "Ce paragraphe est le dernier d'une cellule de tableau : suppression refusee (vider son texte avec replace_paragraph)."
            );
          }
        }
        parent.removeChild(p);
        report.push("delete_paragraph : paragraphe supprime");
        break;
      }
    }
  }

  assertSane(doc, hadFinalSectPr);
  zip.file("word/document.xml", serialize(doc));
  const out = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  return { buffer: out, report };
}
