import { NextResponse } from "next/server";
import { DRIVE_SCOPE, getBaseUrl, isAllowedRedirectUri, oauthError } from "@/lib/oauth-config";
import { nowSeconds, randomToken, s256, seal } from "@/lib/seal";

export const runtime = "nodejs";

/** Duree de vie du state relaye (le temps pour l'utilisateur de consentir). */
const STATE_TTL_SECONDS = 10 * 60;

/**
 * Endpoint /authorize : point d'entree du flow OAuth cote client MCP.
 * On relaie vers l'ecran de consentement Google avec le client_id et le
 * redirect_uri de NOTRE application Google Cloud.
 *
 * Securite :
 * - le redirect_uri du client MCP doit figurer dans ALLOWED_REDIRECT_URIS
 *   (sinon n'importe qui pourrait se faire envoyer le code d'un utilisateur) ;
 * - PKCE S256 est obligatoire cote client MCP, et verifie dans /oauth/token ;
 * - le contexte du client (redirect_uri, state, code_challenge) est scelle
 *   (AES-GCM) dans le state envoye a Google : ni lisible, ni falsifiable ;
 * - on utilise aussi PKCE entre ce serveur et Google (verifier dans le state scelle).
 */
export async function GET(req: Request) {
  const url = new URL(req.url);
  const params = url.searchParams;

  const clientRedirectUri = params.get("redirect_uri") ?? "";
  const clientState = params.get("state") ?? "";
  const clientId = params.get("client_id") ?? "";
  const responseType = params.get("response_type") ?? "code";
  const codeChallenge = params.get("code_challenge") ?? "";
  const codeChallengeMethod = params.get("code_challenge_method") ?? "";

  // Aucune redirection tant que le redirect_uri n'est pas valide (RFC 6749 §4.1.2.1).
  if (!clientRedirectUri || !isAllowedRedirectUri(clientRedirectUri)) {
    return oauthError("invalid_request", "redirect_uri absent ou non autorise");
  }

  const fail = (error: string, description: string) => {
    const back = new URL(clientRedirectUri);
    back.searchParams.set("error", error);
    back.searchParams.set("error_description", description);
    if (clientState) back.searchParams.set("state", clientState);
    return NextResponse.redirect(back.toString());
  };

  if (responseType !== "code") {
    return fail("unsupported_response_type", "Seul response_type=code est supporte");
  }
  if (codeChallengeMethod !== "S256" || !/^[A-Za-z0-9_-]{43}$/.test(codeChallenge)) {
    return fail("invalid_request", "PKCE S256 obligatoire (code_challenge + code_challenge_method=S256)");
  }
  if (clientState.length > 1024 || clientId.length > 512) {
    return fail("invalid_request", "Parametres trop longs");
  }

  const googleClientId = process.env.GOOGLE_OAUTH_CLIENT_ID;
  if (!googleClientId || !process.env.OAUTH_TOKEN_SECRET) {
    return oauthError("server_error", "Serveur OAuth non configure", 500);
  }

  const googleVerifier = randomToken(48);
  const relayState = seal("state", {
    cs: clientState,
    ru: clientRedirectUri,
    cc: codeChallenge,
    cid: clientId,
    gv: googleVerifier,
    exp: nowSeconds() + STATE_TTL_SECONDS,
  });

  const googleAuthUrl = new URL("https://accounts.google.com/o/oauth2/v2/auth");
  googleAuthUrl.searchParams.set("client_id", googleClientId);
  googleAuthUrl.searchParams.set("redirect_uri", `${getBaseUrl(req)}/oauth/callback`);
  googleAuthUrl.searchParams.set("response_type", "code");
  googleAuthUrl.searchParams.set("scope", DRIVE_SCOPE);
  googleAuthUrl.searchParams.set("access_type", "offline");
  googleAuthUrl.searchParams.set("prompt", "consent");
  googleAuthUrl.searchParams.set("state", relayState);
  googleAuthUrl.searchParams.set("code_challenge", s256(googleVerifier));
  googleAuthUrl.searchParams.set("code_challenge_method", "S256");

  return NextResponse.redirect(googleAuthUrl.toString());
}
