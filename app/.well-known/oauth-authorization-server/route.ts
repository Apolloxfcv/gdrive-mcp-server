import { NextResponse } from "next/server";

/**
 * Metadonnees de decouverte OAuth pour le protocole d'autorisation MCP.
 * Perplexity (et d'autres clients MCP) lisent cet endpoint en premier pour
 * decouvrir ou envoyer les requetes /authorize et /token.
 *
 * Ce serveur agit comme un PROXY OAuth : il expose ses propres endpoints
 * /authorize et /oauth/token, qui relaient en interne vers Google
 * (accounts.google.com). Aucun etat n'est garde cote serveur entre les
 * deux etapes du flow autre que ce qui transite dans le "code" et le
 * "state" eux-memes (via des parametres d'URL signes/opaques, pas une
 * session serveur).
 */
export async function GET(req: Request) {
  const baseUrl = new URL(req.url).origin;

  return NextResponse.json({
    issuer: baseUrl,
    authorization_endpoint: `${baseUrl}/authorize`,
    token_endpoint: `${baseUrl}/oauth/token`,
    response_types_supported: ["code"],
    grant_types_supported: ["authorization_code", "refresh_token"],
    code_challenge_methods_supported: ["S256"],
    token_endpoint_auth_methods_supported: ["client_secret_post", "none"],
    scopes_supported: ["https://www.googleapis.com/auth/drive"],
  });
}
