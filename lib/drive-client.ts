import { google } from "googleapis";

/**
 * Construit un client Google Drive authentifie a partir d'un access token
 * fourni par requete (mode stateless). Aucun token n'est jamais stocke
 * en memoire ou sur disque cote serveur.
 */
export function getDriveClient(accessToken: string) {
  const auth = new google.auth.OAuth2();
  auth.setCredentials({ access_token: accessToken });
  return google.drive({ version: "v3", auth });
}

/**
 * Extrait le bearer token Google depuis les headers de la requete HTTP
 * entrante. mcp-handler expose les headers originaux via `extra.request`
 * / `extra.requestInfo` selon la version du SDK; on lit directement
 * l'objet Headers standard Web API par robustesse.
 */
export function extractBearerToken(headers: Headers | undefined): string {
  const raw = headers?.get("authorization") ?? headers?.get("Authorization");
  if (!raw) {
    throw new Error(
      "Aucun token trouve. Le client MCP doit envoyer 'Authorization: Bearer <google_access_token>' a chaque requete."
    );
  }
  const match = raw.match(/^Bearer\s+(.+)$/i);
  if (!match) {
    throw new Error("Le header Authorization doit etre au format 'Bearer <token>'.");
  }
  return match[1];
}
