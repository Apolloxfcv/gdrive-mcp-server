import { createMcpHandler } from "mcp-handler";
import { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import { getDriveClient, extractBearerToken } from "@/lib/drive-client";
import {
  readXlsxAsJson,
  updateXlsxCells,
  createXlsx,
  readDocxAsText,
  createDocx,
  patchDocxPlaceholders,
} from "@/lib/office-utils";

/**
 * Serveur MCP Google Drive - 100% stateless.
 *
 * Aucune session, aucun credential, aucun fichier n'est conserve entre
 * deux requetes. `createMcpHandler` appelle la factory ci-dessous une
 * fois PAR REQUETE HTTP (McpServerFactory), en lui passant un
 * McpRequestContext qui contient `requestInfo` (la Request originale).
 * On extrait le bearer token Google depuis ses headers ICI, une seule
 * fois par requete, puis on le capture en closure dans chaque outil.
 * Aucun etat n'est partage entre deux requetes : une nouvelle instance
 * de McpServer et un nouveau token sont utilises a chaque appel.
 */
const handler = createMcpHandler((ctx) => {
  const server = new McpServer({ name: "gdrive-mcp-server", version: "1.1.0" });

  const token = extractBearerToken(ctx.requestInfo?.headers);
  const drive = getDriveClient(token);

  server.registerTool(
    "drive_list_files",
    {
      description:
        "Liste les fichiers et dossiers Google Drive de l'utilisateur, avec filtre optionnel par requete de recherche Drive (syntaxe q= de l'API Drive) et par dossier parent.",
      inputSchema: z.object({
        query: z
          .string()
          .optional()
          .describe("Requete de recherche Drive, ex: \"name contains 'rapport'\""),
        folderId: z
          .string()
          .optional()
          .describe("ID du dossier parent dans lequel chercher"),
        pageSize: z.number().min(1).max(100).default(20),
      }),
    },
    async ({ query, folderId, pageSize }) => {
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
          { type: "text", text: JSON.stringify(res.data.files ?? [], null, 2) },
        ],
      };
    }
  );

  server.registerTool(
    "drive_read_file",
    {
      description:
        "Lit le contenu texte d'un fichier Google Drive (fichiers Google Docs/Sheets exportes en texte brut, ou fichiers texte bruts).",
      inputSchema: z.object({
        fileId: z.string().describe("ID du fichier Google Drive a lire"),
      }),
    },
    async ({ fileId }) => {
      const meta = await drive.files.get({ fileId, fields: "mimeType, name" });

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

      return { content: [{ type: "text", text }] };
    }
  );

  server.registerTool(
    "drive_create_file",
    {
      description:
        "Cree un nouveau fichier dans Google Drive avec le contenu texte fourni (ecriture).",
      inputSchema: z.object({
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
      }),
    },
    async ({ name, content, mimeType, parentFolderId }) => {
      const res = await drive.files.create({
        requestBody: { name, parents: parentFolderId ? [parentFolderId] : undefined },
        media: { mimeType, body: content },
        fields: "id, name, webViewLink",
      });

      return { content: [{ type: "text", text: JSON.stringify(res.data, null, 2) }] };
    }
  );

  server.registerTool(
    "drive_update_file",
    {
      description: "Met a jour le contenu d'un fichier Google Drive existant (ecriture).",
      inputSchema: z.object({
        fileId: z.string().describe("ID du fichier a mettre a jour"),
        content: z.string().describe("Nouveau contenu texte du fichier"),
        mimeType: z.string().default("text/plain"),
      }),
    },
    async ({ fileId, content, mimeType }) => {
      const res = await drive.files.update({
        fileId,
        media: { mimeType, body: content },
        fields: "id, name, modifiedTime, webViewLink",
      });

      return { content: [{ type: "text", text: JSON.stringify(res.data, null, 2) }] };
    }
  );

  server.registerTool(
    "drive_delete_file",
    {
      description: "Supprime definitivement un fichier Google Drive (ecriture destructive).",
      inputSchema: z.object({
        fileId: z.string().describe("ID du fichier a supprimer"),
      }),
    },
    async ({ fileId }) => {
      await drive.files.delete({ fileId });
      return { content: [{ type: "text", text: `Fichier ${fileId} supprime.` }] };
    }
  );

  server.registerTool(
    "drive_create_folder",
    {
      description: "Cree un nouveau dossier dans Google Drive.",
      inputSchema: z.object({
        name: z.string().describe("Nom du dossier a creer"),
        parentFolderId: z.string().optional(),
      }),
    },
    async ({ name, parentFolderId }) => {
      const res = await drive.files.create({
        requestBody: {
          name,
          mimeType: "application/vnd.google-apps.folder",
          parents: parentFolderId ? [parentFolderId] : undefined,
        },
        fields: "id, name, webViewLink",
      });

      return { content: [{ type: "text", text: JSON.stringify(res.data, null, 2) }] };
    }
  );

  server.registerTool(
    "drive_read_office_file",
    {
      description:
        "Lit le contenu d'un fichier .xlsx ou .docx stocke tel quel sur Drive (pas un Google Sheets/Docs natif). Pour .xlsx, retourne toutes les feuilles sous forme de tableaux JSON. Pour .docx, retourne le texte brut extrait.",
      inputSchema: z.object({
        fileId: z.string().describe("ID du fichier .xlsx ou .docx sur Drive"),
      }),
    },
    async ({ fileId }) => {
      const meta = await drive.files.get({ fileId, fields: "name, mimeType" });
      const name = meta.data.name ?? "";
      const downloaded = await drive.files.get(
        { fileId, alt: "media" },
        { responseType: "arraybuffer" }
      );
      const buffer = Buffer.from(downloaded.data as ArrayBuffer);

      if (name.endsWith(".xlsx")) {
        const sheets = await readXlsxAsJson(buffer);
        return { content: [{ type: "text", text: JSON.stringify(sheets, null, 2) }] };
      }
      if (name.endsWith(".docx")) {
        const text = await readDocxAsText(buffer);
        return { content: [{ type: "text", text }] };
      }
      throw new Error("Ce fichier n'est ni un .xlsx ni un .docx (extension non reconnue).");
    }
  );

  server.registerTool(
    "drive_update_xlsx_cells",
    {
      description:
        "Modifie des cellules specifiques dans un fichier .xlsx existant sur Drive, puis re-uploade le fichier complet (ecriture). Ne fonctionne que sur des fichiers .xlsx binaires, pas sur des Google Sheets natifs.",
      inputSchema: z.object({
        fileId: z.string().describe("ID du fichier .xlsx sur Drive"),
        sheetName: z.string().describe("Nom de la feuille a modifier"),
        updates: z
          .array(
            z.object({
              cell: z.string().describe("Reference de cellule, ex: 'B3'"),
              value: z.union([z.string(), z.number(), z.boolean()]),
            })
          )
          .describe("Liste des cellules a mettre a jour"),
      }),
    },
    async ({ fileId, sheetName, updates }) => {
      const downloaded = await drive.files.get(
        { fileId, alt: "media" },
        { responseType: "arraybuffer" }
      );
      const original = Buffer.from(downloaded.data as ArrayBuffer);
      const updated = await updateXlsxCells(original, sheetName, updates);

      const res = await drive.files.update({
        fileId,
        media: {
          mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          body: updated,
        },
        fields: "id, name, modifiedTime, webViewLink",
      });

      return { content: [{ type: "text", text: JSON.stringify(res.data, null, 2) }] };
    }
  );

  server.registerTool(
    "drive_create_xlsx",
    {
      description: "Cree un nouveau fichier .xlsx sur Drive a partir d'un tableau de lignes (ecriture).",
      inputSchema: z.object({
        name: z.string().describe("Nom du fichier, ex: 'rapport.xlsx'"),
        sheetName: z.string().default("Sheet1"),
        rows: z
          .array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
          .describe("Tableau de lignes, chaque ligne est un tableau de valeurs de cellules"),
        parentFolderId: z.string().optional(),
      }),
    },
    async ({ name, sheetName, rows, parentFolderId }) => {
      const buffer = await createXlsx(sheetName, rows);

      const res = await drive.files.create({
        requestBody: {
          name: name.endsWith(".xlsx") ? name : `${name}.xlsx`,
          parents: parentFolderId ? [parentFolderId] : undefined,
        },
        media: {
          mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          body: buffer,
        },
        fields: "id, name, webViewLink",
      });

      return { content: [{ type: "text", text: JSON.stringify(res.data, null, 2) }] };
    }
  );

  server.registerTool(
    "drive_create_docx",
    {
      description: "Cree un nouveau fichier .docx sur Drive a partir d'une liste de paragraphes (ecriture).",
      inputSchema: z.object({
        name: z.string().describe("Nom du fichier, ex: 'note.docx'"),
        paragraphs: z.array(z.string()).describe("Liste des paragraphes du document"),
        parentFolderId: z.string().optional(),
      }),
    },
    async ({ name, paragraphs, parentFolderId }) => {
      const buffer = await createDocx(paragraphs);

      const res = await drive.files.create({
        requestBody: {
          name: name.endsWith(".docx") ? name : `${name}.docx`,
          parents: parentFolderId ? [parentFolderId] : undefined,
        },
        media: {
          mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          body: buffer,
        },
        fields: "id, name, webViewLink",
      });

      return { content: [{ type: "text", text: JSON.stringify(res.data, null, 2) }] };
    }
  );

  server.registerTool(
    "drive_patch_docx_placeholders",
    {
      description:
        "Remplace des placeholders {{cle}} dans un .docx existant sur Drive par des valeurs, en preservant la mise en forme, puis re-uploade le fichier (ecriture).",
      inputSchema: z.object({
        fileId: z.string().describe("ID du fichier .docx sur Drive"),
        replacements: z
          .record(z.string(), z.string())
          .describe("Map cle -> valeur, ex: {\"nom\": \"Jean Dupont\"} remplace {{nom}}"),
      }),
    },
    async ({ fileId, replacements }) => {
      const downloaded = await drive.files.get(
        { fileId, alt: "media" },
        { responseType: "arraybuffer" }
      );
      const original = Buffer.from(downloaded.data as ArrayBuffer);
      const patched = await patchDocxPlaceholders(original, replacements);

      const res = await drive.files.update({
        fileId,
        media: {
          mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          body: patched,
        },
        fields: "id, name, modifiedTime, webViewLink",
      });

      return { content: [{ type: "text", text: JSON.stringify(res.data, null, 2) }] };
    }
  );

  server.registerTool(
    "drive_rename_file",
    {
      description: "Renomme un fichier ou un dossier Google Drive (ecriture).",
      inputSchema: z.object({
        fileId: z.string().describe("ID du fichier ou dossier"),
        newName: z.string().describe("Nouveau nom"),
      }),
    },
    async ({ fileId, newName }) => {
      const res = await drive.files.update({
        fileId,
        requestBody: { name: newName },
        fields: "id, name, webViewLink",
      });

      return { content: [{ type: "text", text: JSON.stringify(res.data, null, 2) }] };
    }
  );

  server.registerTool(
    "drive_move_file",
    {
      description: "Deplace un fichier ou un dossier vers un nouveau dossier parent sur Google Drive (ecriture).",
      inputSchema: z.object({
        fileId: z.string().describe("ID du fichier ou dossier a deplacer"),
        newParentFolderId: z.string().describe("ID du nouveau dossier parent"),
      }),
    },
    async ({ fileId, newParentFolderId }) => {
      const current = await drive.files.get({ fileId, fields: "parents" });
      const previousParents = (current.data.parents ?? []).join(",");

      const res = await drive.files.update({
        fileId,
        addParents: newParentFolderId,
        removeParents: previousParents,
        fields: "id, name, parents, webViewLink",
      });

      return { content: [{ type: "text", text: JSON.stringify(res.data, null, 2) }] };
    }
  );

  return server;
});

export { handler as GET, handler as POST, handler as DELETE };
