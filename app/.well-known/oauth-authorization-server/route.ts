import { NextResponse } from "next/server";
import { DRIVE_SCOPE, getBaseUrl } from "@/lib/oauth-config";

/**
 * Metadonnees de decouverte OAuth (RFC 8414) pour le protocole
 * d'autorisation MCP. Ce serveur agit comme un PROXY OAuth vers Google :
 * voir /authorize, /oauth/callback et /oauth/token.
 */
export async function GET(req: Request) {
  const baseUrl = getBaseUrl(req);

  return NextResponse.json({
    issuer: baseUrl,
    authorization_endpoint: `${baseUrl}/authorize`,
    token_endpoint: `${baseUrl}/oauth/token`,
    response_types_supported: ["code"],
    grant_types_supported: ["authorization_code", "refresh_token"],
    code_challenge_methods_supported: ["S256"],
    // Clients publics : l'authentification du client repose sur PKCE.
    token_endpoint_auth_methods_supported: ["none"],
    scopes_supported: [DRIVE_SCOPE],
  });
}
