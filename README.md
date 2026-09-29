# gdrive-mcp-server (stateless, Vercel-ready)

Serveur MCP Google Drive **100% stateless**, deploye sur Vercel via des
fonctions serverless Next.js, avec transport **Streamable HTTP**.

## Principe stateless

- Aucun token, aucune session, aucun fichier de credentials n'est jamais
  ecrit ou garde en memoire cote serveur.
- Le client MCP (Perplexity, Claude, Cursor...) doit envoyer le token
  d'acces Google OAuth dans le header `Authorization: Bearer <token>`
  a **chaque** appel d'outil.
- Chaque invocation cree un client Google Drive ephemere avec ce token,
  execute l'appel, puis le jette.

## Outils exposes

| Outil | Action | Ecriture |
|---|---|---|
| `drive_list_files` | Liste/recherche des fichiers | Non |
| `drive_read_file` | Lit le contenu texte (Google Docs/Sheets natifs ou texte brut) | Non |
| `drive_create_file` | Cree un fichier texte brut | Oui |
| `drive_update_file` | Modifie le contenu d'un fichier texte brut | Oui |
| `drive_delete_file` | Met un fichier a la corbeille (recuperable 30 jours) | Oui |
| `drive_create_folder` | Cree un dossier | Oui |
| `drive_rename_file` | Renomme un fichier ou dossier | Oui |
| `drive_move_file` | Deplace un fichier ou dossier vers un autre parent | Oui |
| `drive_read_office_file` | Lit un .xlsx (toutes feuilles en JSON) ou .docx (texte extrait) | Non |
| `drive_update_xlsx_cells` | Modifie des cellules precises (Google Sheets natif : en place via l'API Sheets ; .xlsx : patch XML cible) | Oui |
| `drive_create_xlsx` | Cree un nouveau .xlsx a partir de lignes de donnees | Oui |
| `drive_create_docx` | Cree un nouveau .docx a partir de paragraphes | Oui |
| `drive_patch_docx_placeholders` | Remplace des {{placeholders}} dans un .docx en gardant la mise en forme | Oui |

### Limites connues sur les fichiers Office (.xlsx/.docx)

- Un .xlsx/.docx est un binaire ZIP+XML et l'API Drive n'offre pas d'ecriture
  partielle : le fichier est toujours re-uploade en entier (meme ID, meme
  historique de versions).
- `drive_update_xlsx_cells` limite les degats : sur un .xlsx binaire, seul le XML
  de la feuille visee est patche (cellules modifiees uniquement) ; graphiques,
  images, tableaux croises, mises en forme, plages nommees et formules voisines
  ne sont ni relus ni reecrits. Un controle md5 juste avant l'upload refuse
  d'ecraser un fichier modifie entre-temps. Les formules partagees/matricielles
  ne sont pas modifiables (refus explicite). Les valeurs sont ecrites telles
  quelles (une chaine `=...` n'est pas interpretee comme formule).
- Sur un **Google Sheets natif**, `drive_update_xlsx_cells` utilise l'API Google
  Sheets (`spreadsheets.values.batchUpdate`) : edition directe en place, sans
  telechargement ni upload. L'**API Google Sheets doit etre activee** dans le
  projet Google Cloud (le scope `drive` suffit).
- Les autres outils (lecture, creation, docx) ne fonctionnent que sur des
  .xlsx/.docx binaires uploades tels quels sur Drive.
- Pas de gestion de commentaires ou de suggestions Word/Docs (mode revision)
  dans cette version.

## Authentification (proxy OAuth vers Google)

Le serveur expose ses propres endpoints OAuth (`/authorize`, `/oauth/callback`,
`/oauth/token`, decouverte sous `/.well-known/`) et relaie vers Google, sans
aucun stockage : state, codes et tokens remis au client MCP sont **scelles**
(AES-256-GCM avec `OAUTH_TOKEN_SECRET`). Le client ne voit jamais les tokens
Google en clair, et le endpoint MCP n'accepte que les tokens emis par ce serveur.

- PKCE S256 obligatoire.
- Seuls les `redirect_uri` listes dans `ALLOWED_REDIRECT_URIS` sont acceptes.

## Deploiement sur Vercel

1. Cree un projet Google Cloud, active l'API Google Drive, cree des
   credentials OAuth 2.0 (type "Web application").
2. Ajoute `https://<ton-projet>.vercel.app/oauth/callback` dans les
   "Authorized redirect URIs" de ce credential Google.
3. Renseigne les variables d'environnement (voir `.env.example`) :
   `GOOGLE_OAUTH_CLIENT_ID`, `GOOGLE_OAUTH_CLIENT_SECRET`,
   `OAUTH_TOKEN_SECRET`, `ALLOWED_REDIRECT_URIS`, `PUBLIC_BASE_URL`.
4. Deploie. L'URL du serveur MCP sera :
   `https://<ton-projet>.vercel.app/api/mcp`
5. Ajoute cette URL comme connecteur MCP distant dans ton client.

## Developpement local

```bash
npm install
npm run dev
```

## Securite

Voir [SECURITY_AUDIT.md](SECURITY_AUDIT.md) pour l'audit complet et la
procedure de migration.

- Ne jamais logger le contenu du header `Authorization`.
- Changer `OAUTH_TOKEN_SECRET` revoque tous les tokens emis.
- Limite les scopes OAuth demandes au strict necessaire.
