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
  execute l'appel, puis le jette. Ca rend le serveur trivialement
  scalable horizontalement sur l'infra serverless de Vercel (pas besoin
  de Redis, pas de session affinity).

## Outils exposes

| Outil | Action | Ecriture |
|---|---|---|
| `drive_list_files` | Liste/recherche des fichiers | Non |
| `drive_read_file` | Lit le contenu d'un fichier | Non |
| `drive_create_file` | Cree un fichier | Oui |
| `drive_update_file` | Modifie le contenu d'un fichier | Oui |
| `drive_delete_file` | Supprime un fichier | Oui |
| `drive_create_folder` | Cree un dossier | Oui |

## Obtenir un token Google OAuth

Ce serveur ne gere pas le flow OAuth lui-meme (il resterait sinon avec
etat). Deux options :

1. **Cote client MCP** : si ton client (ex: Claude Desktop, Cursor)
   supporte OAuth 2.1 pour les connecteurs MCP distants, il gerera le
   flow d'autorisation et injectera automatiquement le bearer token
   Google a chaque requete.
2. **Manuellement pour tester** : utilise [Google OAuth Playground]
   (https://developers.google.com/oauthplayground) avec le scope
   `https://www.googleapis.com/auth/drive` pour generer un access token
   de test, et passe-le en header `Authorization: Bearer <token>`.

## Deploiement sur Vercel

1. Cree un projet Google Cloud, active l'API Google Drive, cree des
   credentials OAuth 2.0 (type "Web application").
2. Ajoute l'URL de callback de ton client MCP dans les "Authorized
   redirect URIs" de ce credential Google.
3. Push ce repo sur GitHub, importe-le dans Vercel.
4. Deploie. L'URL du serveur MCP sera :
   `https://<ton-projet>.vercel.app/api/mcp`
5. Ajoute cette URL comme connecteur MCP distant dans Perplexity
   (Account settings -> Connectors -> + Custom connector -> Remote),
   avec authentification OAuth ou Bearer token selon ce que supporte
   le client.

## Developpement local

```bash
npm install
npm run dev
# Le serveur ecoute sur http://localhost:3000/api/mcp
```

## Securite

- Ne jamais logger le contenu du header `Authorization`.
- `drive_delete_file` est destructif et irreversible: implemente une
  confirmation cote client avant d'appeler cet outil.
- Limite les scopes OAuth demandes au strict necessaire
  (`drive.file` plutot que `drive` si tu ne veux acceder qu'aux
  fichiers crees par l'app).
