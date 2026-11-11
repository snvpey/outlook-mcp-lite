import os from 'node:os';
import path from 'node:path';

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// Permissions demandées à Entra ID. Volontairement minimales :
// - pas de Mail.Send : même un token volé ne peut pas envoyer de mail
// - pas de *.Shared : uniquement ta propre boîte
// - pas de Files/Sites : pas d'accès OneDrive ni SharePoint
// MSAL ajoute tout seul openid, profile et offline_access.
export const SCOPES = [
  'User.Read',
  'Mail.ReadWrite',            // lire, créer des brouillons, déplacer, catégoriser
  'MailboxSettings.ReadWrite', // créer des catégories (liste maître)
];

export const LIMITS = {
  maxPageSize: 50,       // messages max par appel de liste
  maxBodyChars: 20000,   // corps de mail tronqué au-delà
  maxBatch: 25,          // messages max par opération de tri groupée
  maxSearchLength: 200,  // longueur max d'une requête de recherche
  requestTimeoutMs: 30000,
  loginTimeoutMs: 5 * 60 * 1000,
};

export function loadConfig(env = process.env) {
  const clientId = (env.AZURE_CLIENT_ID || '').trim();
  const tenantId = (env.AZURE_TENANT_ID || '').trim();

  if (!GUID.test(clientId)) {
    throw new Error('AZURE_CLIENT_ID manquant ou invalide (un GUID est attendu).');
  }
  // On exige le GUID du tenant de l'entreprise. Les valeurs "common" ou
  // "organizations" sont refusées pour ne jamais se connecter à un autre tenant.
  if (!GUID.test(tenantId)) {
    throw new Error('AZURE_TENANT_ID manquant ou invalide (le GUID du tenant est attendu, pas "common").');
  }

  const cacheDir = (env.OUTLOOK_MCP_CACHE_DIR || '').trim()
    || path.join(os.homedir(), '.outlook-mcp-lite');

  return {
    clientId,
    tenantId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    cacheDir,
    // Si "1", le serveur refuse de démarrer quand le stockage chiffré des
    // tokens n'est pas disponible (au lieu de garder les tokens en mémoire).
    requirePersistence: env.OUTLOOK_MCP_REQUIRE_PERSISTENCE === '1',
  };
}
