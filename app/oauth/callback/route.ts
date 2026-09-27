import { NextResponse } from "next/server";
import { isAllowedRedirectUri, oauthError } from "@/lib/oauth-config";
import { nowSeconds, seal, unseal } from "@/lib/seal";

export const runtime = "nodejs";

/** Duree de vie du code d'autorisation remis au client MCP. */
const CODE_TTL_SECONDS = 5 * 60;

type RelayState = { cs: string; ru: string; cc: string; cid: string; gv: string };

/**
 * Endpoint /oauth/callback : recoit le retour de Google apres consentement.
 *
 * Le state est dechiffre et verifie (authenticite + expiration). Le code
 * Google n'est jamais remis tel quel au client : il est scelle avec le
 * code_challenge, le redirect_uri et le client_id du client MCP, ce qui
 * permet a /oauth/token de verifier PKCE sans aucun stockage serveur.
 */
export async function GET(req: Request) {
  const url = new URL(req.url);
  const googleCode = url.searchParams.get("code");
  const googleError = url.searchParams.get("error");

  const relay = unseal<RelayState>("state", url.searchParams.get("state"));
  // Double verification de la liste blanche (au cas ou elle aurait change depuis /authorize).
  if (!relay || !isAllowedRedirectUri(relay.ru)) {
    return oauthError("invalid_request", "state invalide ou expire");
  }

  const back = new URL(relay.ru);
  if (relay.cs) back.searchParams.set("state", relay.cs);

  if (googleError || !googleCode) {
    back.searchParams.set("error", googleError === "access_denied" ? "access_denied" : "server_error");
    return NextResponse.redirect(back.toString());
  }

  const code = seal("code", {
    gc: googleCode,
    gv: relay.gv,
    cc: relay.cc,
    ru: relay.ru,
    cid: relay.cid,
    exp: nowSeconds() + CODE_TTL_SECONDS,
  });
  back.searchParams.set("code", code);

  return NextResponse.redirect(back.toString());
}
