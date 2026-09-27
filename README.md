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
| `drive_delete_file` | Supprime un fichier | Oui |
| `drive_create_folder` | Cree un dossier | Oui |
| `drive_rename_file` | Renomme un fichier ou dossier | Oui |
| `drive_move_file` | Deplace un fichier ou dossier vers un autre parent | Oui |
| `drive_read_office_file` | Lit un .xlsx (toutes feuilles en JSON) ou .docx (texte extrait) | Non |
| `drive_update_xlsx_cells` | Modifie des cellules precises dans un .xlsx existant | Oui |
| `drive_create_xlsx` | Cree un nouveau .xlsx a partir de lignes de donnees | Oui |
| `drive_create_docx` | Cree un nouveau .docx a partir de paragraphes | Oui |
| `drive_patch_docx_placeholders` | Remplace des {{placeholders}} dans un .docx en gardant la mise en forme | Oui |

### Limites connues sur les fichiers Office (.xlsx/.docx)

- Un .xlsx/.docx est un binaire ZIP+XML : toute modification (sauf le patch
  de placeholders) telecharge le fichier complet, le modifie en memoire,
  puis re-uploade le fichier entier. Il n'y a pas d'edition incrementale
  possible cote API Drive pour ces formats.
- Ces outils ne fonctionnent QUE sur des .xlsx/.docx binaires uploades tels
  quels sur Drive. Pour des **Google Sheets/Docs natifs** (crees directement
  dans Drive, pas uploades), il faudrait des outils separes bases sur les
  API Google Sheets et Google Docs, non couverts par ce serveur.
- Pas de gestion de commentaires ou de suggestions Word/Docs (mode revision)
  dans cette version.

## Obtenir un token Google OAuth

Ce serveur ne gere pas le flow OAuth lui-meme (il resterait sinon avec
etat). Deux options :

1. **Cote client MCP** : si ton client supporte OAuth 2.1 pour les
   connecteurs MCP distants, il gerera le flow et injectera automatiquement
   le bearer token Google a chaque requete.
2. **Manuellement pour tester** : utilise Google OAuth Playground
   (https://developers.google.com/oauthplayground) avec le scope
   `https://www.googleapis.com/auth/drive`.

## Deploiement sur Vercel

1. Cree un projet Google Cloud, active l'API Google Drive, cree des
   credentials OAuth 2.0 (type "Web application").
2. Ajoute l'URL de callback de ton client MCP dans les "Authorized
   redirect URIs" de ce credential Google.
3. Push ce repo sur GitHub, importe-le dans Vercel.
4. Deploie. L'URL du serveur MCP sera :
   `https://<ton-projet>.vercel.app/api/mcp`
5. Ajoute cette URL comme connecteur MCP distant dans Perplexity.

## Developpement local

```bash
npm install
npm run dev
```

## Securite

- Ne jamais logger le contenu du header `Authorization`.
- `drive_delete_file` est destructif et irreversible.
- Limite les scopes OAuth demandes au strict necessaire.
