import { NextResponse } from "next/server";
import { DRIVE_SCOPE, getBaseUrl } from "@/lib/oauth-config";

/**
 * Metadonnees de ressource protegee (RFC 9728). mcp-handler pointe vers
 * cette URL dans le header WWW-Authenticate des reponses 401 ; elle
 * repondait 404 auparavant.
 */
export async function GET(req: Request) {
  const baseUrl = getBaseUrl(req);

  return NextResponse.json({
    resource: `${baseUrl}/api/mcp`,
    authorization_servers: [baseUrl],
    scopes_supported: [DRIVE_SCOPE],
    bearer_methods_supported: ["header"],
  });
}
