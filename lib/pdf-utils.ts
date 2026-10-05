import { extractText, getDocumentProxy } from "unpdf";
import { marked } from "marked";
import { UserFacingError } from "./tool-guards";

/** Extrait le texte d'un PDF (couche texte uniquement, pas d'OCR), page par page. */
export async function readPdfPages(buffer: Buffer) {
  if (buffer.subarray(0, 5).toString("latin1") !== "%PDF-") {
    throw new UserFacingError("Ce fichier n'est pas un PDF valide.");
  }
  try {
    const pdf = await getDocumentProxy(new Uint8Array(buffer));
    const { totalPages, text } = await extractText(pdf, { mergePages: false });
    return { totalPages, pages: text as string[] };
  } catch {
    throw new UserFacingError("PDF illisible (corrompu, protege par mot de passe ou format non supporte).");
  }
}

/** Markdown -> HTML complet, utilise comme source de conversion vers un Google Doc natif. */
export function markdownToHtml(markdown: string): string {
  const body = marked.parse(markdown, { async: false }) as string;
  return `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${body}</body></html>`;
}
