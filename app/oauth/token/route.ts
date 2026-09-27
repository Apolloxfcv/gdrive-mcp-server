import { NextResponse } from "next/server";

/**
 * Endpoint /oauth/token : echange le code d'autorisation Google contre un
 * access token + refresh token Google, et les retourne au client MCP
 * (Perplexity) au format standard OAuth token response.
 *
 * Stateless : chaque appel fait un echange direct avec Google
 * (oauth2.googleapis.com/token) sans jamais persister le resultat cote
 * serveur. Le client MCP est responsable de stocker le token retourne.
 */
export async function POST(req: Request) {
  const contentType = req.headers.get("content-type") ?? "";
  let params: URLSearchParams;

  if (contentType.includes("application/json")) {
    const body = await req.json();
    params = new URLSearchParams(body);
  } else {
    const body = await req.text();
    params = new URLSearchParams(body);
  }

  const grantType = params.get("grant_type");
  const googleClientId = process.env.GOOGLE_OAUTH_CLIENT_ID;
  const googleClientSecret = process.env.GOOGLE_OAUTH_CLIENT_SECRET;

  if (!googleClientId || !googleClientSecret) {
    return NextResponse.json(
      { error: "server_error", error_description: "Credentials Google non configures" },
      { status: 500 }
    );
  }

  const baseUrl = new URL(req.url).origin;
  const ourCallbackUri = `${baseUrl}/oauth/callback`;

  const googleParams = new URLSearchParams();
  googleParams.set("client_id", googleClientId);
  googleParams.set("client_secret", googleClientSecret);

  if (grantType === "authorization_code") {
    const code = params.get("code");
    if (!code) {
      return NextResponse.json({ error: "invalid_request", error_description: "code manquant" }, { status: 400 });
    }
    googleParams.set("grant_type", "authorization_code");
    googleParams.set("code", code);
    googleParams.set("redirect_uri", ourCallbackUri);
  } else if (grantType === "refresh_token") {
    const refreshToken = params.get("refresh_token");
    if (!refreshToken) {
      return NextResponse.json({ error: "invalid_request", error_description: "refresh_token manquant" }, { status: 400 });
    }
    googleParams.set("grant_type", "refresh_token");
    googleParams.set("refresh_token", refreshToken);
  } else {
    return NextResponse.json({ error: "unsupported_grant_type" }, { status: 400 });
  }

  const googleRes = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: googleParams.toString(),
  });

  const googleData = await googleRes.json();

  if (!googleRes.ok) {
    return NextResponse.json(
      { error: "invalid_grant", error_description: googleData.error_description ?? "Echange de token Google echoue" },
      { status: 400 }
    );
  }

  return NextResponse.json({
    access_token: googleData.access_token,
    token_type: googleData.token_type ?? "Bearer",
    expires_in: googleData.expires_in,
    refresh_token: googleData.refresh_token,
    scope: googleData.scope,
  });
}
