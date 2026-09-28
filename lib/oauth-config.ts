import { NextResponse } from "next/server";

export const DRIVE_SCOPE = "https://www.googleapis.com/auth/drive";

/**
 * URL publique du serveur. PUBLIC_BASE_URL est recommande en production :
 * sans lui, on retombe sur l'origine de la requete (comportement historique).
 * Le callback Google `${base}/oauth/callback` doit etre declare a l'identique
 * dans la console Google Cloud.
 */
export function getBaseUrl(req: Request): string {
  const configured = process.env.PUBLIC_BASE_URL?.trim().replace(/\/+$/, "");
  return configured || new URL(req.url).origin;
}

/**
 * Liste blanche des redirect_uri des clients MCP (comparaison exacte).
 * Ex: ALLOWED_REDIRECT_URIS="https://claude.ai/api/mcp/auth_callback,https://www.perplexity.ai/..."
 * Liste vide = aucun client autorise (fail closed).
 */
export function getAllowedRedirectUris(): string[] {
  return (process.env.ALLOWED_REDIRECT_URIS ?? "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
}

export function isAllowedRedirectUri(uri: string): boolean {
  return getAllowedRedirectUris().includes(uri);
}

/**
 * Periode de transition : accepte encore les tokens Google "bruts" emis par
 * l'ancienne version (avant scellement), apres verification de leur audience
 * aupres de Google. A desactiver des que les clients se sont reconnectes.
 */
export function legacyTokensAllowed(): boolean {
  return process.env.ALLOW_LEGACY_GOOGLE_TOKENS === "true";
}

export function getGoogleCredentials() {
  const clientId = process.env.GOOGLE_OAUTH_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_OAUTH_CLIENT_SECRET;
  return clientId && clientSecret ? { clientId, clientSecret } : undefined;
}

/** Reponse d'erreur OAuth, jamais mise en cache (RFC 6749 §5.1/5.2). */
export function oauthError(error: string, description: string, status = 400) {
  return NextResponse.json(
    { error, error_description: description },
    { status, headers: NO_STORE_HEADERS }
  );
}

export const NO_STORE_HEADERS = {
  "Cache-Control": "no-store",
  Pragma: "no-cache",
};
