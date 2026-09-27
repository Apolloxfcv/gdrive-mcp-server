import { NextResponse } from "next/server";
import {
  NO_STORE_HEADERS,
  getBaseUrl,
  getGoogleCredentials,
  legacyTokensAllowed,
  oauthError,
} from "@/lib/oauth-config";
import { isSealed, nowSeconds, s256, safeEqual, seal, unseal } from "@/lib/seal";

export const runtime = "nodejs";

const MAX_BODY_BYTES = 16 * 1024;

type SealedCode = { gc: string; gv: string; cc: string; ru: string; cid: string };
type SealedRefresh = { grt: string; cid: string };

type GoogleTokenResponse = {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  scope?: string;
  error?: string;
};

/**
 * Endpoint /oauth/token.
 *
 * - authorization_code : dechiffre le code emis par /oauth/callback, verifie
 *   PKCE (code_verifier), le redirect_uri et le client_id, puis echange le
 *   code aupres de Google.
 * - refresh_token : n'accepte que des refresh tokens scelles par ce serveur
 *   (sauf periode de transition ALLOW_LEGACY_GOOGLE_TOKENS).
 *
 * Les tokens Google ne sortent jamais en clair : le client MCP recoit des
 * tokens scelles, que seul ce serveur sait ouvrir. Un token Google vole
 * ailleurs n'est donc pas utilisable ici, et un token de ce serveur n'est
 * pas utilisable directement contre les API Google.
 */
export async function POST(req: Request) {
  const params = await readParams(req);
  if (!params) return oauthError("invalid_request", "Corps de requete invalide");

  const google = getGoogleCredentials();
  if (!google || !process.env.OAUTH_TOKEN_SECRET) {
    return oauthError("server_error", "Serveur OAuth non configure", 500);
  }

  const grantType = params.get("grant_type");
  const clientId = params.get("client_id") ?? "";

  if (grantType === "authorization_code") {
    const code = unseal<SealedCode>("code", params.get("code"));
    if (!code) return oauthError("invalid_grant", "code invalide ou expire");

    const verifier = params.get("code_verifier") ?? "";
    if (!/^[A-Za-z0-9._~-]{43,128}$/.test(verifier) || !safeEqual(s256(verifier), code.cc)) {
      return oauthError("invalid_grant", "code_verifier invalide");
    }
    const redirectUri = params.get("redirect_uri");
    if (redirectUri !== null && redirectUri !== code.ru) {
      return oauthError("invalid_grant", "redirect_uri different de celui de l'autorisation");
    }
    if (code.cid && clientId && clientId !== code.cid) {
      return oauthError("invalid_grant", "client_id different de celui de l'autorisation");
    }

    const data = await callGoogle(google, {
      grant_type: "authorization_code",
      code: code.gc,
      code_verifier: code.gv,
      redirect_uri: `${getBaseUrl(req)}/oauth/callback`,
    });
    if (!data?.access_token) return oauthError("invalid_grant", "Echange de token Google echoue");

    return tokenResponse(data, code.cid, data.refresh_token);
  }

  if (grantType === "refresh_token") {
    const raw = params.get("refresh_token") ?? "";
    let googleRefresh: string | undefined;
    let boundClientId = clientId;

    if (isSealed(raw)) {
      const rt = unseal<SealedRefresh>("refresh_token", raw);
      if (!rt) return oauthError("invalid_grant", "refresh_token invalide");
      if (rt.cid && clientId && clientId !== rt.cid) {
        return oauthError("invalid_grant", "client_id different de celui du refresh_token");
      }
      googleRefresh = rt.grt;
      boundClientId = rt.cid;
    } else if (raw && legacyTokensAllowed()) {
      // Transition : ancien refresh token Google brut, re-emis scelle ci-dessous.
      googleRefresh = raw;
    } else {
      return oauthError("invalid_grant", "refresh_token invalide");
    }

    const data = await callGoogle(google, { grant_type: "refresh_token", refresh_token: googleRefresh });
    if (!data?.access_token) return oauthError("invalid_grant", "Rafraichissement Google echoue");

    // Google ne renvoie generalement pas de nouveau refresh token : on garde l'actuel.
    return tokenResponse(data, boundClientId, data.refresh_token ?? googleRefresh);
  }

  return oauthError("unsupported_grant_type", "grant_type non supporte");
}

async function readParams(req: Request): Promise<URLSearchParams | undefined> {
  try {
    const text = await req.text();
    if (text.length > MAX_BODY_BYTES) return undefined;
    const contentType = req.headers.get("content-type") ?? "";
    if (contentType.includes("application/json")) {
      const body = JSON.parse(text);
      if (!body || typeof body !== "object" || Array.isArray(body)) return undefined;
      const params = new URLSearchParams();
      for (const [k, v] of Object.entries(body)) {
        if (typeof v === "string") params.set(k, v);
      }
      return params;
    }
    return new URLSearchParams(text);
  } catch {
    return undefined;
  }
}

async function callGoogle(
  creds: { clientId: string; clientSecret: string },
  params: Record<string, string>
): Promise<GoogleTokenResponse | undefined> {
  try {
    const res = await fetch("https://oauth2.googleapis.com/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: creds.clientId,
        client_secret: creds.clientSecret,
        ...params,
      }).toString(),
      cache: "no-store",
    });
    const data = (await res.json()) as GoogleTokenResponse;
    if (!res.ok) {
      // On ne logge que le code d'erreur Google, jamais les parametres (code / tokens).
      console.warn("Google token endpoint error:", res.status, data?.error);
      return undefined;
    }
    return data;
  } catch {
    return undefined;
  }
}

function tokenResponse(data: GoogleTokenResponse, clientId: string, googleRefresh?: string) {
  const expiresIn = Math.max(60, Math.min(data.expires_in ?? 3600, 3600));
  return NextResponse.json(
    {
      access_token: seal("access_token", {
        gat: data.access_token,
        cid: clientId,
        exp: nowSeconds() + expiresIn,
      }),
      token_type: "Bearer",
      expires_in: expiresIn,
      refresh_token: googleRefresh
        ? seal("refresh_token", { grt: googleRefresh, cid: clientId })
        : undefined,
      scope: data.scope,
    },
    { headers: NO_STORE_HEADERS }
  );
}
