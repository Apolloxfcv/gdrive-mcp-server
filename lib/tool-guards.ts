import { z } from "zod";

/**
 * Garde-fous communs aux outils MCP : validation des entrees, limites de
 * taille (memoire d'une fonction serverless) et erreurs assainies.
 */

export const MAX_DOWNLOAD_BYTES = Number(process.env.MAX_DOWNLOAD_BYTES) || 20 * 1024 * 1024;
export const MAX_UNCOMPRESSED_BYTES = 5 * MAX_DOWNLOAD_BYTES;
export const MAX_TEXT_OUTPUT_CHARS = 500_000;

/**
 * ID Drive : caracteres alphanumeriques, '-' et '_' (ou l'alias "root").
 * Empeche l'injection dans la syntaxe q= de l'API Drive.
 */
export const driveId = z
  .string()
  .regex(/^[A-Za-z0-9_-]{1,256}$/, "ID Drive invalide");

export const fileName = z.string().min(1).max(255);

export const mimeType = z
  .string()
  .regex(/^[a-z0-9][a-z0-9.+-]*\/[a-z0-9][a-z0-9.+-]*$/i, "Type MIME invalide")
  .max(127);

export const cellRef = z.string().regex(/^[A-Z]{1,3}[1-9][0-9]{0,6}$/i, "Reference de cellule invalide");

export function truncate(text: string): string {
  if (text.length <= MAX_TEXT_OUTPUT_CHARS) return text;
  return (
    text.slice(0, MAX_TEXT_OUTPUT_CHARS) +
    `\n\n[... tronque : ${text.length - MAX_TEXT_OUTPUT_CHARS} caracteres supplementaires non affiches]`
  );
}

export function assertDownloadable(size: string | null | undefined) {
  if (size && Number(size) > MAX_DOWNLOAD_BYTES) {
    throw new UserFacingError(
      `Fichier trop volumineux (${Math.round(Number(size) / 1024 / 1024)} Mo, limite ${Math.round(
        MAX_DOWNLOAD_BYTES / 1024 / 1024
      )} Mo).`
    );
  }
}

/** Options gaxios limitant la taille de la reponse telechargee. */
export const downloadLimits = { maxContentLength: MAX_DOWNLOAD_BYTES } as const;

export class UserFacingError extends Error {}

type ToolResult = { content: { type: "text"; text: string }[]; isError?: boolean };

/**
 * Enveloppe un handler d'outil : toute erreur est convertie en message
 * court, sans stack, sans configuration de requete (qui contient le
 * header Authorization dans les erreurs gaxios).
 */
export function guarded<A, E>(fn: (args: A, extra: E) => Promise<ToolResult>) {
  return async (args: A, extra: E): Promise<ToolResult> => {
    try {
      return await fn(args, extra);
    } catch (err) {
      return { isError: true, content: [{ type: "text", text: describeError(err) }] };
    }
  };
}

function describeError(err: unknown): string {
  if (err instanceof UserFacingError) return err.message;
  const e = err as {
    response?: { status?: number; data?: { error?: { message?: string } } };
    message?: string;
  };
  if (e?.response?.status) {
    const status = e.response.status;
    const apiMessage = e.response.data?.error?.message;
    return `Erreur Google Drive (${status})${apiMessage ? ` : ${String(apiMessage).slice(0, 300)}` : ""}`;
  }
  console.error("Tool error:", e?.message ?? "unknown");
  return "Erreur interne lors de l'execution de l'outil.";
}
