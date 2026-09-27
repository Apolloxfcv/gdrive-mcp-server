import { createMcpHandler } from "mcp-handler";
import { z } from "zod";
import { getDriveClient, extractBearerToken } from "@/lib/drive-client";

/**
 * Serveur MCP Google Drive - 100% stateless.
 *
 * Aucune session, aucun credential, aucun fichier n'est conserve entre
 * deux requetes. Chaque appel d'outil recupere le token Google depuis le
 * header Authorization de la requete HTTP en cours, cree un client Drive
 * ephemere, execute l'appel API, puis jette le client. Ce modele est
 * compatible avec les fonctions serverless Vercel (aucun etat partage
 * requis entre invocations, donc scalable horizontalement sans backend
 * de session comme Redis).
 */
const handler = createMcpHandler(
  (server) => {
    server.tool(
      "drive_list_files",
      "Liste les fichiers et dossiers Google Drive de l'utilisateur, avec filtre optionnel par requete de recherche Drive (syntaxe q= de l'API Drive) et par dossier parent.",
      {
        query: z
          .string()
          .optional()
          .describe("Requete de recherche Drive, ex: \"name contains 'rapport'\""),
        folderId: z
          .string()
          .optional()
          .describe("ID du dossier parent dans lequel chercher"),
        pageSize: z.number().min(1).max(100).default(20),
      },
      async ({ query, folderId, pageSize }, extra) => {
        const token = extractBearerToken(extra?.requestInfo?.headers as Headers | undefined);
        const drive = getDriveClient(token);

        const qParts: string[] = ["trashed = false"];
        if (query) qParts.push(query);
        if (folderId) qParts.push(`'${folderId}' in parents`);

        const res = await drive.files.list({
          q: qParts.join(" and "),
          pageSize,
          fields: "files(id, name, mimeType, modifiedTime, size, webViewLink)",
        });

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(res.data.files ?? [], null, 2),
            },
          ],
        };
      }
    );

    server.tool(
      "drive_read_file",
      "Lit le contenu texte d'un fichier Google Drive (fichiers Google Docs/Sheets exportes en texte brut, ou fichiers texte bruts).",
      {
        fileId: z.string().describe("ID du fichier Google Drive a lire"),
      },
      async ({ fileId }, extra) => {
        const token = extractBearerToken(extra?.requestInfo?.headers as Headers | undefined);
        const drive = getDriveClient(token);

        const meta = await drive.files.get({
          fileId,
          fields: "mimeType, name",
        });

        let text: string;
        if (meta.data.mimeType?.startsWith("application/vnd.google-apps")) {
          const exported = await drive.files.export(
            { fileId, mimeType: "text/plain" },
            { responseType: "text" }
          );
          text = exported.data as unknown as string;
        } else {
          const downloaded = await drive.files.get(
            { fileId, alt: "media" },
            { responseType: "text" }
          );
          text = downloaded.data as unknown as string;
        }

        return {
          content: [{ type: "text", text }],
        };
      }
    );

    server.tool(
      "drive_create_file",
      "Cree un nouveau fichier dans Google Drive avec le contenu texte fourni (ecriture).",
      {
        name: z.string().describe("Nom du fichier a creer"),
        content: z.string().describe("Contenu texte du fichier"),
        mimeType: z
          .string()
          .default("text/plain")
          .describe("Type MIME du fichier, ex: text/plain, text/markdown, application/json"),
        parentFolderId: z
          .string()
          .optional()
          .describe("ID du dossier parent ou creer le fichier"),
      },
      async ({ name, content, mimeType, parentFolderId }, extra) => {
        const token = extractBearerToken(extra?.requestInfo?.headers as Headers | undefined);
        const drive = getDriveClient(token);

        const res = await drive.files.create({
          requestBody: {
            name,
            parents: parentFolderId ? [parentFolderId] : undefined,
          },
          media: {
            mimeType,
            body: content,
          },
          fields: "id, name, webViewLink",
        });

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(res.data, null, 2),
            },
          ],
        };
      }
    );

    server.tool(
      "drive_update_file",
      "Met a jour le contenu d'un fichier Google Drive existant (ecriture).",
      {
        fileId: z.string().describe("ID du fichier a mettre a jour"),
        content: z.string().describe("Nouveau contenu texte du fichier"),
        mimeType: z.string().default("text/plain"),
      },
      async ({ fileId, content, mimeType }, extra) => {
        const token = extractBearerToken(extra?.requestInfo?.headers as Headers | undefined);
        const drive = getDriveClient(token);

        const res = await drive.files.update({
          fileId,
          media: {
            mimeType,
            body: content,
          },
          fields: "id, name, modifiedTime, webViewLink",
        });

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(res.data, null, 2),
            },
          ],
        };
      }
    );

    server.tool(
      "drive_delete_file",
      "Supprime definitivement un fichier Google Drive (ecriture destructive).",
      {
        fileId: z.string().describe("ID du fichier a supprimer"),
      },
      async ({ fileId }, extra) => {
        const token = extractBearerToken(extra?.requestInfo?.headers as Headers | undefined);
        const drive = getDriveClient(token);

        await drive.files.delete({ fileId });

        return {
          content: [
            { type: "text", text: `Fichier ${fileId} supprime.` },
          ],
        };
      }
    );

    server.tool(
      "drive_create_folder",
      "Cree un nouveau dossier dans Google Drive.",
      {
        name: z.string().describe("Nom du dossier a creer"),
        parentFolderId: z.string().optional(),
      },
      async ({ name, parentFolderId }, extra) => {
        const token = extractBearerToken(extra?.requestInfo?.headers as Headers | undefined);
        const drive = getDriveClient(token);

        const res = await drive.files.create({
          requestBody: {
            name,
            mimeType: "application/vnd.google-apps.folder",
            parents: parentFolderId ? [parentFolderId] : undefined,
          },
          fields: "id, name, webViewLink",
        });

        return {
          content: [
            { type: "text", text: JSON.stringify(res.data, null, 2) },
          ],
        };
      }
    );
  },
  {
    // Aucune capacite de session/streaming persistante n'est declaree :
    // le serveur reste 100% stateless, chaque requete HTTP est independante.
  },
  {
    basePath: "/api",
    verboseLogs: false,
  }
);

export { handler as GET, handler as POST, handler as DELETE };
