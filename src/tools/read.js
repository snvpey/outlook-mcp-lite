import { z } from 'zod';
import { LIMITS } from '../config.js';
import { seg } from '../graph.js';
import { folderSegment } from '../folders.js';
import { messageSummary, messageFull, folderSummary, ok } from '../format.js';
import { messageId, folderRef } from './schemas.js';

const SUMMARY_FIELDS = 'id,receivedDateTime,from,subject,isRead,isDraft,flag,categories,hasAttachments,bodyPreview';
const FULL_FIELDS = 'id,conversationId,parentFolderId,receivedDateTime,from,toRecipients,ccRecipients,'
  + 'subject,body,isRead,isDraft,flag,categories,hasAttachments';

const READ_ONLY = { readOnlyHint: true, openWorldHint: false };

/** Nettoie une requête de recherche : pas de guillemets, pas de caractères de contrôle. */
export function cleanSearchQuery(query) {
  return query.replace(/[\u0000-\u001f"\\]/g, ' ').trim();
}

export function registerReadTools(server, { graph, wrap }) {
  server.registerTool('outlook_list_folders', {
    title: 'Lister les dossiers',
    description: 'Liste les dossiers mail (racine par défaut, ou les sous-dossiers d\'un dossier).',
    inputSchema: {
      parentFolder: folderRef.optional(),
    },
    annotations: READ_ONLY,
  }, wrap(async ({ parentFolder }) => {
    const path = parentFolder
      ? `/me/mailFolders/${folderSegment(parentFolder)}/childFolders`
      : '/me/mailFolders';
    const data = await graph.get(path, {
      query: {
        $top: 100,
        $select: 'id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount',
      },
    });
    return ok((data.value || []).map(folderSummary));
  }));

  server.registerTool('outlook_list_messages', {
    title: 'Lister les mails d\'un dossier',
    description: 'Liste les mails d\'un dossier, du plus récent au plus ancien. Renvoie un résumé (pas le corps complet).',
    inputSchema: {
      folder: folderRef.default('inbox'),
      top: z.number().int().min(1).max(LIMITS.maxPageSize).default(20),
      skip: z.number().int().min(0).max(5000).default(0),
      unreadOnly: z.boolean().default(false),
    },
    annotations: READ_ONLY,
  }, wrap(async ({ folder, top, skip, unreadOnly }) => {
    const query = {
      $select: SUMMARY_FIELDS,
      $top: top,
      $skip: skip,
      $orderby: 'receivedDateTime desc',
    };
    if (unreadOnly) {
      // Graph exige que le champ du $orderby apparaisse aussi dans le $filter.
      query.$filter = 'receivedDateTime ge 1900-01-01T00:00:00Z and isRead eq false';
    }
    const data = await graph.get(`/me/mailFolders/${folderSegment(folder)}/messages`, { query });
    return ok({
      folder,
      count: data.value?.length || 0,
      nextSkip: data['@odata.nextLink'] ? skip + top : null,
      messages: (data.value || []).map(messageSummary),
    });
  }));

  server.registerTool('outlook_search_messages', {
    title: 'Rechercher des mails',
    description: 'Recherche dans toute la boîte (syntaxe KQL possible, ex: from:dupont subject:facture). '
      + 'Résultats triés par pertinence.',
    inputSchema: {
      query: z.string().min(1).max(LIMITS.maxSearchLength),
      top: z.number().int().min(1).max(LIMITS.maxPageSize).default(20),
    },
    annotations: READ_ONLY,
  }, wrap(async ({ query, top }) => {
    const cleaned = cleanSearchQuery(query);
    if (!cleaned) throw new Error('Requête de recherche vide après nettoyage.');
    const data = await graph.get('/me/messages', {
      query: { $search: `"${cleaned}"`, $select: SUMMARY_FIELDS, $top: top },
    });
    return ok({ query: cleaned, count: data.value?.length || 0, messages: (data.value || []).map(messageSummary) });
  }));

  server.registerTool('outlook_get_message', {
    title: 'Lire un mail',
    description: 'Lit un mail complet (corps en texte brut) et la liste de ses pièces jointes (nom, type, taille). '
      + 'Le contenu des pièces jointes n\'est pas lu.',
    inputSchema: {
      messageId,
      maxChars: z.number().int().min(500).max(LIMITS.maxBodyChars).default(LIMITS.maxBodyChars),
    },
    annotations: READ_ONLY,
  }, wrap(async ({ messageId: id, maxChars }) => {
    const message = await graph.get(`/me/messages/${seg(id)}`, {
      query: { $select: FULL_FIELDS },
      // Graph convertit le HTML en texte : pas besoin de parser du HTML ici.
      headers: { Prefer: 'outlook.body-content-type="text"' },
    });
    let attachments = [];
    if (message.hasAttachments) {
      const data = await graph.get(`/me/messages/${seg(id)}/attachments`, {
        query: { $select: 'id,name,contentType,size,isInline' },
      });
      attachments = data.value || [];
    }
    return ok(messageFull(message, attachments, maxChars));
  }));

  server.registerTool('outlook_list_categories', {
    title: 'Lister les catégories',
    description: 'Liste les catégories (labels de couleur) disponibles dans la boîte.',
    inputSchema: {},
    annotations: READ_ONLY,
  }, wrap(async () => {
    const data = await graph.get('/me/outlook/masterCategories');
    return ok((data.value || []).map((c) => ({ displayName: c.displayName, color: c.color })));
  }));
}
