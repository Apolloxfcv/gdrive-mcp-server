import {
  createCipheriv,
  createDecipheriv,
  createHash,
  hkdfSync,
  randomBytes,
  timingSafeEqual,
} from "node:crypto";

/**
 * Chiffrement authentifie (AES-256-GCM) des objets qui transitent par le
 * client ou le navigateur : state OAuth relaye, code d'autorisation,
 * access token et refresh token emis par ce serveur.
 *
 * C'est ce qui permet de rester stateless SANS faire confiance au contenu
 * renvoye par le client : un blob scelle ne peut etre ni lu, ni forge, ni
 * modifie sans OAUTH_TOKEN_SECRET. Chaque type de blob a son propre
 * "purpose" (lie en AAD) pour empecher de presenter un state comme un
 * code, un code comme un access token, etc.
 *
 * Faire tourner OAUTH_TOKEN_SECRET revoque instantanement tous les tokens
 * emis (les clients devront se reconnecter).
 */

export type SealPurpose = "state" | "code" | "access_token" | "refresh_token";

const PREFIX = "v1.";
let cachedKey: Buffer | undefined;

function getKey(): Buffer {
  if (cachedKey) return cachedKey;
  const secret = process.env.OAUTH_TOKEN_SECRET;
  if (!secret || secret.length < 32) {
    throw new Error("OAUTH_TOKEN_SECRET manquant ou trop court (32 caracteres minimum).");
  }
  cachedKey = Buffer.from(hkdfSync("sha256", secret, "gdrive-mcp-server", "seal-v1", 32));
  return cachedKey;
}

export function seal(purpose: SealPurpose, payload: Record<string, unknown>): string {
  const iv = randomBytes(12);
  const cipher = createCipheriv("aes-256-gcm", getKey(), iv);
  cipher.setAAD(Buffer.from(purpose));
  const ciphertext = Buffer.concat([cipher.update(JSON.stringify(payload), "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return PREFIX + Buffer.concat([iv, tag, ciphertext]).toString("base64url");
}

/**
 * Retourne le payload si le blob est authentique, du bon type et non
 * expire (champ `exp` en secondes epoch, optionnel). Sinon `undefined`.
 */
export function unseal<T extends Record<string, unknown>>(
  purpose: SealPurpose,
  value: string | null | undefined
): T | undefined {
  if (!value || !value.startsWith(PREFIX) || value.length > 8192) return undefined;
  try {
    const raw = Buffer.from(value.slice(PREFIX.length), "base64url");
    if (raw.length < 12 + 16 + 1) return undefined;
    const decipher = createDecipheriv("aes-256-gcm", getKey(), raw.subarray(0, 12));
    decipher.setAAD(Buffer.from(purpose));
    decipher.setAuthTag(raw.subarray(12, 28));
    const plaintext = Buffer.concat([decipher.update(raw.subarray(28)), decipher.final()]);
    const payload = JSON.parse(plaintext.toString("utf8")) as T;
    const exp = payload.exp;
    if (typeof exp === "number" && exp < nowSeconds()) return undefined;
    return payload;
  } catch {
    return undefined;
  }
}

export function isSealed(value: string | undefined): boolean {
  return !!value && value.startsWith(PREFIX);
}

export function nowSeconds(): number {
  return Math.floor(Date.now() / 1000);
}

export function randomToken(bytes = 32): string {
  return randomBytes(bytes).toString("base64url");
}

/** code_challenge S256 = BASE64URL(SHA256(code_verifier)) (RFC 7636). */
export function s256(verifier: string): string {
  return createHash("sha256").update(verifier).digest("base64url");
}

export function safeEqual(a: string, b: string): boolean {
  const ab = Buffer.from(a);
  const bb = Buffer.from(b);
  return ab.length === bb.length && timingSafeEqual(ab, bb);
}
