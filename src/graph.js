import { LIMITS } from './config.js';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

/**
 * Liste blanche des appels Graph autorisés. Tout le reste est refusé AVANT
 * l'envoi de la requête. En particulier :
 * - aucune méthode DELETE
 * - pas de /send, /forward, /reply, /replyAll (seulement createReply* qui fait un brouillon)
 * - pas de /permanentDelete, pas de /copy
 * - uniquement /me (pas d'autres boîtes)
 * Les ids sont encodés avec encodeURIComponent, donc un segment ne contient jamais "/".
 */
const SEG = '[^/]+';
export const ALLOWED_ROUTES = [
  ['GET', new RegExp(`^/me$`)],
  ['GET', new RegExp(`^/me/mailFolders$`)],
  ['GET', new RegExp(`^/me/mailFolders/${SEG}$`)],
  ['GET', new RegExp(`^/me/mailFolders/${SEG}/(childFolders|messages)$`)],
  ['POST', new RegExp(`^/me/mailFolders$`)],
  ['POST', new RegExp(`^/me/mailFolders/${SEG}/childFolders$`)],
  ['GET', new RegExp(`^/me/messages$`)],
  ['GET', new RegExp(`^/me/messages/${SEG}$`)],
  ['GET', new RegExp(`^/me/messages/${SEG}/attachments$`)],
  ['POST', new RegExp(`^/me/messages$`)], // crée un brouillon
  ['PATCH', new RegExp(`^/me/messages/${SEG}$`)],
  ['POST', new RegExp(`^/me/messages/${SEG}/(createReply|createReplyAll|move)$`)],
  ['GET', new RegExp(`^/me/outlook/masterCategories$`)],
  ['POST', new RegExp(`^/me/outlook/masterCategories$`)],
];

// Champs modifiables par PATCH sur un message.
export const ALLOWED_PATCH_FIELDS = new Set([
  'isRead', 'flag', 'categories',
  'subject', 'body', 'toRecipients', 'ccRecipients', 'bccRecipients', // brouillons uniquement
]);

export class GraphError extends Error {
  constructor(status, code, message) {
    super(`Graph ${status}${code ? ` ${code}` : ''} : ${message || 'erreur inconnue'}`);
    this.status = status;
    this.code = code;
  }
}

export function assertAllowed(method, path, body) {
  const ok = ALLOWED_ROUTES.some(([m, re]) => m === method && re.test(path));
  if (!ok) {
    throw new Error(`Appel Graph refusé par la liste blanche : ${method} ${path}`);
  }
  if (method === 'PATCH' && body) {
    const bad = Object.keys(body).filter((k) => !ALLOWED_PATCH_FIELDS.has(k));
    if (bad.length) {
      throw new Error(`Champs non autorisés dans la modification : ${bad.join(', ')}`);
    }
  }
}

/** Encode un id ou un nom de dossier pour l'utiliser comme segment d'URL. */
export function seg(value) {
  return encodeURIComponent(value);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export class GraphClient {
  constructor(auth, { fetchImpl = fetch } = {}) {
    this.auth = auth;
    this.fetchImpl = fetchImpl;
  }

  async request(method, path, { query, body, headers } = {}) {
    assertAllowed(method, path, body);

    const url = new URL(GRAPH_BASE + path);
    for (const [key, value] of Object.entries(query || {})) {
      if (value !== undefined && value !== null) url.searchParams.set(key, String(value));
    }

    const maxAttempts = 3;
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      const token = await this.auth.getToken();
      const res = await this.fetchImpl(url, {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          ...(body ? { 'Content-Type': 'application/json' } : {}),
          ...headers,
        },
        body: body ? JSON.stringify(body) : undefined,
        redirect: 'error',
        signal: AbortSignal.timeout(LIMITS.requestTimeoutMs),
      });

      // 429 : requête non traitée, on peut réessayer quelle que soit la méthode.
      // 503/504 : on ne réessaie que les GET pour ne pas créer de doublons.
      const retryable = res.status === 429 || (method === 'GET' && (res.status === 503 || res.status === 504));
      if (retryable && attempt < maxAttempts) {
        const retryAfter = Number(res.headers.get('retry-after'));
        const waitMs = Number.isFinite(retryAfter) && retryAfter > 0
          ? Math.min(retryAfter * 1000, 30000)
          : 1000 * 2 ** attempt;
        await sleep(waitMs);
        continue;
      }

      if (!res.ok) {
        let code;
        let message;
        try {
          const data = await res.json();
          code = data?.error?.code;
          message = data?.error?.message;
        } catch {
          message = res.statusText;
        }
        throw new GraphError(res.status, code, message);
      }

      if (res.status === 204) return null;
      const text = await res.text();
      return text ? JSON.parse(text) : null;
    }
    throw new GraphError(429, 'TooManyRequests', 'trop de requêtes, réessaie plus tard');
  }

  get(path, options) { return this.request('GET', path, options); }
  post(path, body, options) { return this.request('POST', path, { ...options, body }); }
  patch(path, body, options) { return this.request('PATCH', path, { ...options, body }); }
}
