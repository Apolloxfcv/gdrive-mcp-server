import { NextResponse } from "next/server";

/**
 * Endpoint /oauth/callback : recoit le retour de Google apres consentement.
 * On decode le state relaye pour retrouver le redirect_uri et le state
 * ORIGINAUX du client MCP, puis on redirige vers ce client avec le code
 * d'autorisation Google. Le client MCP echangera ensuite ce code contre
 * un token via /oauth/token.
 *
 * Toujours stateless : aucune donnee n'est persistee, tout transite dans
 * les parametres d'URL (le "code" Google est ephemere et a usage unique).
 */
export async function GET(req: Request) {
  const url = new URL(req.url);
  const code = url.searchParams.get("code");
  const relayState = url.searchParams.get("state");
  const error = url.searchParams.get("error");

  if (error) {
    return NextResponse.json({ error, error_description: "Autorisation Google refusee" }, { status: 400 });
  }
  if (!code || !relayState) {
    return NextResponse.json({ error: "invalid_request", error_description: "code ou state manquant" }, { status: 400 });
  }

  let decoded: {
    clientState: string;
    clientRedirectUri: string;
    codeChallenge: string;
    codeChallengeMethod: string;
  };
  try {
    decoded = JSON.parse(Buffer.from(relayState, "base64url").toString("utf-8"));
  } catch {
    return NextResponse.json({ error: "invalid_request", error_description: "state invalide" }, { status: 400 });
  }

  const finalRedirect = new URL(decoded.clientRedirectUri);
  finalRedirect.searchParams.set("code", code);
  if (decoded.clientState) finalRedirect.searchParams.set("state", decoded.clientState);

  return NextResponse.redirect(finalRedirect.toString());
}
