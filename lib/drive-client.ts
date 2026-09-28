import { createHash } from "node:crypto";
import { google } from "googleapis";
import type { AuthInfo } from "@modelcontextprotocol/sdk/server/auth/types.js";
import { DRIVE_SCOPE, legacyTokensAllowed } from "./oauth-config";
import { isSealed, unseal } from "./seal";

/**
 * Construit un client Google Drive authentifie a partir d'un access token
 * Google fourni par requete (mode stateless). Aucun token n'est stocke.
 */
export function getDriveClient(accessToken: string) {
  const auth = new google.auth.OAuth2();
  auth.setCredentials({ access_token: accessToken });
  return google.drive({ version: "v3", auth });
}

type SealedAccess = { gat: string; cid: string; exp: number };

/**
 * Verifie le bearer token presente au endpoint MCP. Seuls les tokens emis
 * (scelles) par /oauth/token sont acceptes : un token Google obtenu par une
 * autre application ne peut pas etre rejoue ici ("token passthrough").
 */
export async function verifyAccessToken(bearerToken?: string): Promise<AuthInfo | undefined> {
  if (!bearerToken) return undefined;

  if (isSealed(bearerToken)) {
    const at = unseal<SealedAccess>("access_token", bearerToken);
    if (!at?.gat) return undefined;
    return {
      token: at.gat,
      clientId: at.cid || "mcp-client",
      scopes: [DRIVE_SCOPE],
      expiresAt: at.exp,
    };
  }

  if (legacyTokensAllowed()) return verifyLegacyGoogleToken(bearerToken);
  return undefined;
}

// ---------- Transition : tokens Google bruts emis par l'ancienne version ----------

const legacyCache = new Map<string, { info: AuthInfo; until: number }>();
const LEGACY_CACHE_MS = 5 * 60 * 1000;

async function verifyLegacyGoogleToken(token: string): Promise<AuthInfo | undefined> {
  const key = createHash("sha256").update(token).digest("hex");
  const cached = legacyCache.get(key);
  if (cached && cached.until > Date.now()) return cached.info;

  try {
    const res = await fetch("https://oauth2.googleapis.com/tokeninfo", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ access_token: token }).toString(),
      cache: "no-store",
    });
    if (!res.ok) return undefined;
    const data = (await res.json()) as { aud?: string; scope?: string; exp?: string };

    // Le token doit avoir ete emis pour NOTRE application Google, avec le scope Drive.
    if (!data.aud || data.aud !== process.env.GOOGLE_OAUTH_CLIENT_ID) return undefined;
    const scopes = (data.scope ?? "").split(" ");
    if (!scopes.includes(DRIVE_SCOPE)) return undefined;

    const info: AuthInfo = {
      token,
      clientId: "legacy-google-token",
      scopes,
      expiresAt: data.exp ? Number(data.exp) : undefined,
    };
    if (legacyCache.size > 500) legacyCache.clear();
    legacyCache.set(key, { info, until: Date.now() + LEGACY_CACHE_MS });
    return info;
  } catch {
    return undefined;
  }
}
