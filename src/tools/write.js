import { z } from 'zod';
import { seg } from '../graph.js';
import { folderSegment } from '../folders.js';
import { ok, toRecipients, sanitizeUntrusted } from '../format.js';
import { messageId, messageIds, folderRef, displayName, emailList, bodyType } from './schemas.js';

// Aucun outil ici n'envoie ni ne supprime quoi que ce soit.
const SAFE_WRITE = { readOnlyHint: false, destructiveHint: false, openWorldHint: false };

const CATEGORY_COLORS = ['none', ...Array.from({ length: 25 }, (_, i) => `preset${i}`)];

const DRAFT_NOTE = 'Brouillon enregistré dans le dossier Brouillons. Il n\'est PAS envoyé : '
  + 'relis-le et envoie-le toi-même depuis Outlook.';

/** Applique fn à chaque id, séquentiellement, et renvoie le résultat de chacun. */
async function forEachMessage(ids, fn) {
  const results = [];
  for (const id of ids) {
    try {
      results.push({ id, ok: true, ...(await fn(id)) });
    } catch (err) {
      results.push({ id, ok: false, error: err.message });
    }
  }
  return {
    succeeded: results.filter((r) => r.ok).length,
    failed: results.filter((r) => !r.ok).length,
    results,
  };
}

function draftSummary(d) {
  return {
    id: d.id,
    subject: sanitizeUntrusted(d.subject, 300).text,
    to: (d.toRecipients || []).map((r) => r.emailAddress?.address),
    isDraft: d.isDraft,
    note: DRAFT_NOTE,
  };
}

export function registerWriteTools(server, { graph, guard, wrap }) {
  // ---------- Brouillons ----------

  server.registerTool('outlook_create_draft', {
    title: 'Créer un brouillon',
    description: 'Crée un nouveau mail en brouillon. Rien n\'est envoyé.',
    inputSchema: {
      subject: z.string().max(500),
      body: z.string().max(100000),
      bodyType,
      to: emailList.default([]),
      cc: emailList.default([]),
      bcc: emailList.default([]),
    },
    annotations: SAFE_WRITE,
  }, wrap(async ({ subject, body, bodyType: type, to, cc, bcc }) => {
    const draft = await graph.post('/me/messages', {
      subject,
      body: { contentType: type === 'html' ? 'HTML' : 'Text', content: body },
      toRecipients: toRecipients(to),
      ccRecipients: toRecipients(cc),
      bccRecipients: toRecipients(bcc),
    });
    return ok(draftSummary(draft));
  }));

  server.registerTool('outlook_create_reply_draft', {
    title: 'Préparer une réponse en brouillon',
    description: 'Crée un brouillon de réponse (ou de réponse à tous) à un mail. Rien n\'est envoyé.',
    inputSchema: {
      messageId,
      comment: z.string().max(100000).describe('Texte de la réponse, ajouté au-dessus du mail cité'),
      replyAll: z.boolean().default(false),
    },
    annotations: SAFE_WRITE,
  }, wrap(async ({ messageId: id, comment, replyAll }) => {
    const action = replyAll ? 'createReplyAll' : 'createReply';
    const draft = await graph.post(`/me/messages/${seg(id)}/${action}`, { comment });
    return ok(draftSummary(draft));
  }));

  server.registerTool('outlook_update_draft', {
    title: 'Modifier un brouillon',
    description: 'Modifie un brouillon existant (objet, corps, destinataires). Refuse si le mail n\'est pas un brouillon.',
    inputSchema: {
      draftId: messageId,
      subject: z.string().max(500).optional(),
      body: z.string().max(100000).optional(),
      bodyType,
      to: emailList.optional(),
      cc: emailList.optional(),
      bcc: emailList.optional(),
    },
    annotations: { ...SAFE_WRITE, idempotentHint: true },
  }, wrap(async ({ draftId, subject, body, bodyType: type, to, cc, bcc }) => {
    const current = await graph.get(`/me/messages/${seg(draftId)}`, { query: { $select: 'id,isDraft' } });
    if (!current.isDraft) {
      throw new Error('Ce mail n\'est pas un brouillon, il ne peut pas être modifié.');
    }
    const patch = {};
    if (subject !== undefined) patch.subject = subject;
    if (body !== undefined) patch.body = { contentType: type === 'html' ? 'HTML' : 'Text', content: body };
    if (to !== undefined) patch.toRecipients = toRecipients(to);
    if (cc !== undefined) patch.ccRecipients = toRecipients(cc);
    if (bcc !== undefined) patch.bccRecipients = toRecipients(bcc);
    if (!Object.keys(patch).length) throw new Error('Aucune modification demandée.');
    const draft = await graph.patch(`/me/messages/${seg(draftId)}`, patch);
    return ok(draftSummary(draft));
  }));

  // ---------- Dossiers ----------

  server.registerTool('outlook_create_folder', {
    title: 'Créer un dossier',
    description: 'Crée un dossier mail, à la racine ou dans un dossier parent.',
    inputSchema: {
      displayName,
      parentFolder: folderRef.optional(),
    },
    annotations: SAFE_WRITE,
  }, wrap(async ({ displayName: name, parentFolder }) => {
    let path = '/me/mailFolders';
    if (parentFolder) {
      await guard.assertAllowedDestination(parentFolder);
      path = `/me/mailFolders/${folderSegment(parentFolder)}/childFolders`;
    }
    const folder = await graph.post(path, { displayName: name });
    return ok({ id: folder.id, displayName: folder.displayName, parentFolderId: folder.parentFolderId });
  }));

  server.registerTool('outlook_move_messages', {
    title: 'Déplacer des mails',
    description: 'Déplace des mails vers un dossier. Les dossiers de suppression (Éléments supprimés et '
      + 'leurs sous-dossiers) sont interdits. Attention : un mail déplacé change d\'id.',
    inputSchema: {
      messageIds,
      destinationFolder: folderRef,
    },
    annotations: SAFE_WRITE,
  }, wrap(async ({ messageIds: ids, destinationFolder }) => {
    const target = await guard.assertAllowedDestination(destinationFolder);
    const summary = await forEachMessage(ids, async (id) => {
      const moved = await graph.post(`/me/messages/${seg(id)}/move`, { destinationId: target.id });
      return { newId: moved.id };
    });
    return ok({ destination: target.displayName, ...summary });
  }));

  // ---------- Catégories ----------

  server.registerTool('outlook_create_category', {
    title: 'Créer une catégorie',
    description: 'Crée une catégorie (label de couleur) utilisable sur les mails.',
    inputSchema: {
      displayName,
      color: z.enum(CATEGORY_COLORS).default('preset0').describe('preset0 à preset24, ou none'),
    },
    annotations: SAFE_WRITE,
  }, wrap(async ({ displayName: name, color }) => {
    const category = await graph.post('/me/outlook/masterCategories', { displayName: name, color });
    return ok({ displayName: category.displayName, color: category.color });
  }));

  server.registerTool('outlook_set_categories', {
    title: 'Ajouter ou retirer des catégories',
    description: 'Ajoute et/ou retire des catégories sur des mails. Les autres catégories du mail sont conservées.',
    inputSchema: {
      messageIds,
      add: z.array(displayName).max(20).default([]),
      remove: z.array(displayName).max(20).default([]),
    },
    annotations: { ...SAFE_WRITE, idempotentHint: true },
  }, wrap(async ({ messageIds: ids, add, remove }) => {
    if (!add.length && !remove.length) throw new Error('Rien à ajouter ni à retirer.');
    const toRemove = new Set(remove.map((c) => c.toLowerCase()));
    const summary = await forEachMessage(ids, async (id) => {
      const m = await graph.get(`/me/messages/${seg(id)}`, { query: { $select: 'id,categories' } });
      const next = (m.categories || []).filter((c) => !toRemove.has(c.toLowerCase()));
      for (const c of add) {
        if (!next.some((x) => x.toLowerCase() === c.toLowerCase())) next.push(c);
      }
      await graph.patch(`/me/messages/${seg(id)}`, { categories: next });
      return { categories: next };
    });
    return ok(summary);
  }));

  // ---------- Lu / non lu, drapeau ----------

  server.registerTool('outlook_mark_read', {
    title: 'Marquer lu / non lu',
    description: 'Marque des mails comme lus ou non lus.',
    inputSchema: {
      messageIds,
      isRead: z.boolean().default(true),
    },
    annotations: { ...SAFE_WRITE, idempotentHint: true },
  }, wrap(async ({ messageIds: ids, isRead }) => {
    const summary = await forEachMessage(ids, async (id) => {
      await graph.patch(`/me/messages/${seg(id)}`, { isRead });
      return {};
    });
    return ok(summary);
  }));

  server.registerTool('outlook_set_flag', {
    title: 'Drapeau de suivi',
    description: 'Met, termine ou retire le drapeau de suivi sur des mails.',
    inputSchema: {
      messageIds,
      status: z.enum(['flagged', 'complete', 'notFlagged']),
    },
    annotations: { ...SAFE_WRITE, idempotentHint: true },
  }, wrap(async ({ messageIds: ids, status }) => {
    const summary = await forEachMessage(ids, async (id) => {
      await graph.patch(`/me/messages/${seg(id)}`, { flag: { flagStatus: status } });
      return {};
    });
    return ok(summary);
  }));
}
