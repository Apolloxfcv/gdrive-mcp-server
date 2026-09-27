# Audit de sécurité — gdrive-mcp-server

- **Date :** 27/09/2026
- **Périmètre :** branche `main` @ `8e214a4` (toutes les routes, `lib/`, dépendances, historique git)
- **Méthode :** revue de code manuelle, `npm audit`, recherche de secrets dans l'historique git, tests dynamiques en local de la version corrigée (sans toucher au déploiement de production)

## Synthèse

| # | Sévérité | Constat | Statut (branche `security-hardening`) |
|---|---|---|---|
| 1 | **Critique** | Aucune liste blanche des `redirect_uri` : vol du code OAuth puis prise de contrôle complète du Drive d'une victime | Corrigé |
| 2 | **Critique** | PKCE annoncé mais jamais vérifié ; `/oauth/token` échange n'importe quel code sans authentifier le client | Corrigé |
| 3 | **Élevée** | `state` relayé non signé (simple base64 JSON), falsifiable, sans expiration | Corrigé |
| 4 | **Élevée** | `/oauth/token` sert de relais public pour le `client_secret` : tout refresh token Google volé est rafraîchissable par n'importe qui, sans limite | Corrigé |
| 5 | **Élevée** | Token passthrough : `/api/mcp` accepte n'importe quel token Google (audience non vérifiée, scopes codés en dur) | Corrigé |
| 6 | Moyenne | `drive_delete_file` supprime **définitivement** (contourne la corbeille) : exposé à la prompt injection | Corrigé (corbeille) |
| 7 | Moyenne | Pas de limite de taille (téléchargement, zip bombs .xlsx/.docx, sortie) : déni de service / coût | Corrigé |
| 8 | Moyenne | Prompt injection indirecte via le contenu des fichiers Drive | Atténué (annotations + avertissement) |
| 9 | Faible | Injection dans la syntaxe `q=` via `folderId` | Corrigé |
| 10 | Faible | Erreurs non assainies, crash 500 sur JSON/URL invalides | Corrigé |
| 11 | Faible | Réponses token sans `Cache-Control: no-store`, pas d'en-têtes de sécurité, `Referer` peut faire fuiter le code | Corrigé |
| 12 | Faible | `/.well-known/oauth-protected-resource` annoncé dans les 401 mais absent (404) | Corrigé |
| 13 | Faible | Dépendances : pas de lockfile ; `npm audit` = 7 vulnérabilités (1 élevée postcss/next au build, uuid via googleapis/exceljs) | Partiellement corrigé |
| 14 | Info | Scope `drive` complet (accès à tout le Drive) | Documenté |
| 15 | Info | Pas de rate limiting sur les endpoints OAuth | Recommandation |

Pas de secret trouvé dans l'historique git (`.env.example` vide, aucun motif `GOCSPX`, `ya29.`, `1//`, `AIza`).

---

## Constats détaillés

### 1. Critique — Open redirect sur `/authorize` → vol de compte Drive

`app/authorize/route.ts` accepte n'importe quel `redirect_uri` et `app/oauth/callback/route.ts` y redirige avec le code Google.

**Scénario :** un attaquant envoie à la victime
`https://<serveur>/authorize?redirect_uri=https://attaquant.tld/x&response_type=code`.
La victime voit le vrai écran de consentement Google de *votre* application, accepte, et son navigateur part vers `https://attaquant.tld/x?code=…`. L'attaquant poste ce code sur `/oauth/token` (aucune authentification client, aucun PKCE) et récupère un **access token + refresh token Google avec le scope `drive` complet** : lecture, modification et suppression de tout le Drive, de façon persistante.

**Correctif :** liste blanche `ALLOWED_REDIRECT_URIS` (comparaison exacte), vérifiée à `/authorize` **et** au callback ; aucune redirection tant que le `redirect_uri` n'est pas validé (RFC 6749 §4.1.2.1).

### 2. Critique — PKCE non vérifié, codes rejouables par un tiers

La découverte annonce `code_challenge_methods_supported: ["S256"]`, `/authorize` stocke `code_challenge`… mais `/oauth/token` ignore `code_verifier`. Le code Google est remis en clair au client, donc quiconque l'intercepte (logs, historique navigateur, `Referer`, extension) peut l'échanger.

**Correctif :** PKCE S256 obligatoire ; le code remis au client est un blob **chiffré et authentifié** (AES-256-GCM) contenant le code Google, le `code_challenge`, le `redirect_uri`, le `client_id` et une expiration de 5 min. `/oauth/token` vérifie `SHA256(code_verifier) == code_challenge` (comparaison à temps constant), le `redirect_uri` et le `client_id`. PKCE est aussi utilisé entre le serveur et Google.

### 3. Élevée — `state` falsifiable

Le commentaire parle de « paramètres signés/opaques », mais le `state` envoyé à Google n'est qu'un JSON en base64. Un attaquant peut le forger (redirect_uri arbitraire, challenge au choix), aucune expiration, et `new URL(decoded.clientRedirectUri)` plante en 500 sur une valeur invalide.

**Correctif :** `state` scellé (AES-GCM, AAD = usage), expiration 10 min.

### 4. Élevée — Relais public du `client_secret`

`/oauth/token` avec `grant_type=refresh_token` ajoute votre `client_secret` à n'importe quel refresh token reçu. Un refresh token Google fuité (qui, seul, est inutilisable sans le secret) devient donc exploitable indéfiniment par n'importe qui via votre serveur.

**Correctif :** les refresh tokens remis aux clients sont scellés par le serveur ; les tokens Google bruts sont refusés (sauf période de transition explicite, voir Migration). Changer `OAUTH_TOKEN_SECRET` révoque tout.

### 5. Élevée — Token passthrough / confused deputy

`verifyToken` accepte n'importe quelle chaîne comme bearer et renvoie des scopes codés en dur : `requiredScopes` ne vérifie rien. Tout token Google avec le scope Drive, émis pour **n'importe quelle** application, est accepté. La spécification MCP interdit explicitement ce « token passthrough ».

**Correctif :** seuls les access tokens scellés par ce serveur sont acceptés (vérification locale, sans appel réseau) ; `expiresAt` est renseigné, donc les tokens expirés sont refusés en 401 par `mcp-handler`.

### 6. Moyenne — Suppression définitive

`drive.files.delete` contourne la corbeille. Combinée à la prompt injection (#8), une seule instruction cachée dans un document lu peut détruire des fichiers de façon irréversible.

**Correctif :** `drive_delete_file` met à la corbeille (récupérable 30 jours) ; annotations `destructiveHint` sur tous les outils qui écrasent des données, pour que le client demande confirmation.

### 7. Moyenne — Absence de limites de ressources

Téléchargement intégral en mémoire sans limite de taille, décompression .xlsx/.docx sans garde-fou (zip bomb : quelques Mo → plusieurs Go), sortie texte non bornée, tableaux d'entrée non bornés.

**Correctif :** limite de téléchargement (20 Mo par défaut, `MAX_DOWNLOAD_BYTES`) vérifiée via les métadonnées et `maxContentLength` ; contrôle de la taille décompressée déclarée avant tout parsing ; troncature de sortie à 500 000 caractères ; bornes zod sur toutes les entrées.

### 8. Moyenne — Prompt injection indirecte

Le contenu des fichiers (Docs, xlsx, docx partagés par des tiers) est renvoyé tel quel au modèle, qui dispose aussi d'outils d'écriture. C'est inhérent aux serveurs MCP et ne peut pas être entièrement réglé côté serveur.

**Atténuation :** descriptions d'outils signalant le contenu comme non fiable, annotations lecture/écriture, suppression réversible. **Recommandation :** garder la confirmation manuelle des outils d'écriture activée côté client.

### 9. Faible — Injection dans la requête Drive

`'${folderId}' in parents` : un `folderId` contenant une apostrophe modifie la requête. Impact limité (même utilisateur, `query` est déjà libre), mais corrigé : tous les IDs sont validés par `^[A-Za-z0-9_-]{1,256}$` et la `query` libre est parenthésée pour ne pas neutraliser `trashed = false`.

### 10. Faible — Gestion des erreurs

Les erreurs gaxios contiennent la configuration de la requête (dont le header `Authorization`) ; selon la sérialisation, un message pouvait exposer des détails. Un JSON invalide sur `/oauth/token` provoquait un 500.
**Correctif :** wrapper `guarded` qui ne renvoie que le statut HTTP et le message de l'API Google ; parsing défensif et taille de corps bornée sur `/oauth/token` ; logs sans paramètres sensibles.

### 11. Faible — En-têtes HTTP

Ajout de `Cache-Control: no-store` / `Pragma: no-cache` sur les réponses token (RFC 6749 §5.1), et globalement `Referrer-Policy: no-referrer`, `X-Content-Type-Options`, `X-Frame-Options`, HSTS, suppression de `X-Powered-By`.

### 12. Faible — Métadonnées de ressource protégée

Ajout de `/.well-known/oauth-protected-resource` (RFC 9728). `token_endpoint_auth_methods_supported` passe à `["none"]` : le serveur n'a jamais vérifié de secret client, l'annonce de `client_secret_post` était trompeuse.

### 13. Faible — Dépendances

- Pas de `package-lock.json` : chaque build Vercel résout des versions différentes (reproductibilité, supply chain). → lockfile ajouté.
- `googleapis` ^144 → ^182 (supprime la dépendance vulnérable `uuid` via `googleapis-common`).
- Restant : `postcss` via `next` 15 (outil de build, non exploitable à l'exécution ici) et `uuid` via `exceljs` (fonction non utilisée). Correction complète = passage à `next` 16, à planifier séparément.

### 14. Info — Scope OAuth

Le scope `https://www.googleapis.com/auth/drive` donne accès à tout le Drive. Nécessaire pour lire des fichiers existants ; si l'usage se limite aux fichiers créés par l'outil, `drive.file` réduirait fortement l'impact d'une compromission.

### 15. Info — Rate limiting

Aucune limite de débit sur `/authorize` et `/oauth/token`. Recommandé : règle Vercel Firewall (ex. 30 requêtes/min/IP sur `/oauth/*`).

---

## Migration (à faire AVANT de merger sur `main`)

La branche est sans effet sur la production tant qu'elle n'est pas mergée (Vercel ne crée qu'un déploiement de preview).

1. Dans Vercel → Settings → Environment Variables (Production), ajouter :
   - `OAUTH_TOKEN_SECRET` : `node -e "console.log(require('crypto').randomBytes(48).toString('base64url'))"`
   - `ALLOWED_REDIRECT_URIS` : l'URL de callback exacte de chaque client MCP (Perplexity, Claude : `https://claude.ai/api/mcp/auth_callback`, …). Pour la trouver : le `redirect_uri` apparaît dans la requête `/authorize` des logs Vercel.
   - `PUBLIC_BASE_URL` : `https://<projet>.vercel.app`
   - `ALLOW_LEGACY_GOOGLE_TOKENS=true` **temporairement** : les clients déjà connectés continuent de fonctionner et reçoivent des tokens scellés à leur prochain refresh.
2. Merger, vérifier qu'une connexion neuve fonctionne.
3. Après quelques jours, retirer `ALLOW_LEGACY_GOOGLE_TOKENS`.
4. Par précaution (constat #1 exploitable depuis la mise en ligne) : vérifier dans https://myaccount.google.com/permissions les accès accordés, et envisager de régénérer le `client_secret` Google.

## Tests réalisés sur la version corrigée (local)

| Test | Résultat |
|---|---|
| `redirect_uri` hors liste blanche | 400, aucune redirection |
| `/authorize` sans PKCE | redirection vers le client avec `error=invalid_request` |
| `state` forgé au callback | 400 `state invalide ou expire` |
| Mauvais `code_verifier` | `invalid_grant` |
| `redirect_uri` différent à l'échange | `invalid_grant` |
| Refresh token Google brut | `invalid_grant` |
| Bearer Google quelconque sur `/api/mcp` | 401 |
| Token scellé valide + `tools/list` | 13 outils, annotations présentes |
| `folderId` avec injection | rejeté par la validation zod |
| Erreur Google (token invalide) | message assaini `Erreur Google Drive (401) : …` |
