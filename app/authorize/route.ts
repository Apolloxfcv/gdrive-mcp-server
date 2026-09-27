import { NextResponse } from "next/server";

/**
 * Endpoint /authorize : point d'entree du flow OAuth cote client MCP
 * (Perplexity). On relaie directement vers l'ecran de consentement Google
 * en reutilisant les memes parametres (client_id, redirect_uri, state,
 * code_challenge, etc.), SAUF que le client_id et le redirect_uri sont
 * ceux de NOTRE application Google Cloud, pas ceux envoyes par le client
 * MCP (qui ne connait que ce serveur, pas Google directement).
 *
 * Stateless : aucune session n'est creee ici. Le "state" fourni par le
 * client MCP est retransmis tel quel a Google, qui nous le renverra sur
 * le callback ; a charge du serveur MCP de le retransmettre au client MCP
 * au retour (fait ici via un pass-through direct, sans stockage).
 */
export async function GET(req: Request) {
  const url = new URL(req.url);
  const clientState = url.searchParams.get("state") ?? "";
  const clientRedirectUri = url.searchParams.get("redirect_uri") ?? "";
  const codeChallenge = url.searchParams.get("code_challenge") ?? "";
  const codeChallengeMethod = url.searchParams.get("code_challenge_method") ?? "S256";

  if (!clientRedirectUri) {
    return NextResponse.json(
      { error: "invalid_request", error_description: "redirect_uri manquant" },
      { status: 400 }
    );
  }

  const googleClientId = process.env.GOOGLE_OAUTH_CLIENT_ID;
  if (!googleClientId) {
    return NextResponse.json(
      { error: "server_error", error_description: "GOOGLE_OAUTH_CLIENT_ID non configure" },
      { status: 500 }
    );
  }

  const baseUrl = url.origin;
  const ourCallbackUri = `${baseUrl}/oauth/callback`;

  // On encode le redirect_uri et le state ORIGINAUX du client MCP dans le
  // parametre state que l'on envoie a Google, pour pouvoir les restituer
  // au moment du callback, sans avoir besoin de les stocker cote serveur.
  const relayState = Buffer.from(
    JSON.stringify({ clientState, clientRedirectUri, codeChallenge, codeChallengeMethod })
  ).toString("base64url");

  const googleAuthUrl = new URL("https://accounts.google.com/o/oauth2/v2/auth");
  googleAuthUrl.searchParams.set("client_id", googleClientId);
  googleAuthUrl.searchParams.set("redirect_uri", ourCallbackUri);
  googleAuthUrl.searchParams.set("response_type", "code");
  googleAuthUrl.searchParams.set("scope", "https://www.googleapis.com/auth/drive");
  googleAuthUrl.searchParams.set("access_type", "offline");
  googleAuthUrl.searchParams.set("prompt", "consent");
  googleAuthUrl.searchParams.set("state", relayState);

  return NextResponse.redirect(googleAuthUrl.toString());
}
