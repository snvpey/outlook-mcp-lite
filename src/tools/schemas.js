import { z } from 'zod';
import { LIMITS } from '../config.js';

// Les ids Graph sont en base64 (lettres, chiffres, - _ = + / .).
const ID_PATTERN = /^[A-Za-z0-9_\-=+/.]+$/;

export const messageId = z.string().min(1).max(512).regex(ID_PATTERN, 'id de message invalide')
  .describe('Id du message (obtenu via outlook_list_messages ou outlook_search_messages)');

export const messageIds = z.array(messageId).min(1).max(LIMITS.maxBatch)
  .describe(`Liste d'ids de messages (max ${LIMITS.maxBatch})`);

export const folderRef = z.string().min(1).max(512).regex(ID_PATTERN, 'référence de dossier invalide')
  .describe('Id du dossier, ou nom standard : inbox, drafts, sentitems, archive, junkemail');

export const displayName = z.string().trim().min(1).max(255)
  .refine((s) => !/[\u0000-\u001f]/.test(s), 'caractères de contrôle interdits');

export const emailList = z.array(z.email()).max(50);

export const bodyType = z.enum(['text', 'html']).default('text')
  .describe('Format du corps : text (défaut) ou html');
