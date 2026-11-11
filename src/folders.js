import { seg, GraphError } from './graph.js';

// Noms de dossiers standards acceptés à la place d'un id.
export const WELL_KNOWN_FOLDERS = [
  'inbox', 'drafts', 'sentitems', 'archive', 'junkemail', 'deleteditems', 'outbox', 'msgfolderroot',
];

// Destinations interdites pour un déplacement ou la création d'un dossier.
// Déplacer un mail dans "Éléments supprimés" revient à le supprimer, donc on bloque.
// Outbox est bloqué pour éviter tout effet de bord lié à l'envoi.
export const FORBIDDEN_DESTINATIONS = [
  'deleteditems',
  'outbox',
  'recoverableitemsroot',
  'recoverableitemsdeletions',
  'recoverableitemspurges',
  'recoverableitemsversions',
];

/** Transforme une référence de dossier (nom standard ou id) en segment d'URL. */
export function folderSegment(ref) {
  const lower = String(ref).toLowerCase();
  return WELL_KNOWN_FOLDERS.includes(lower) ? lower : seg(ref);
}

/**
 * Vérifie qu'un dossier de destination n'est pas un dossier de suppression
 * ni un de ses sous-dossiers (on remonte la chaîne des parents).
 */
export class FolderGuard {
  constructor(graph) {
    this.graph = graph;
    this.forbiddenIds = null;
    this.rootId = null;
  }

  async loadForbiddenIds() {
    if (this.forbiddenIds) return;
    const ids = new Set();
    for (const name of FORBIDDEN_DESTINATIONS) {
      try {
        const folder = await this.graph.get(`/me/mailFolders/${name}`, { query: { $select: 'id' } });
        if (folder?.id) ids.add(folder.id);
      } catch (err) {
        // Certains dossiers "recoverable" n'existent pas ou ne sont pas lisibles.
        if (!(err instanceof GraphError) || (err.status !== 404 && err.status !== 403 && err.status !== 400)) throw err;
      }
    }
    const root = await this.graph.get('/me/mailFolders/msgfolderroot', { query: { $select: 'id' } });
    this.rootId = root?.id || null;
    this.forbiddenIds = ids;
  }

  async assertAllowedDestination(ref) {
    if (FORBIDDEN_DESTINATIONS.includes(String(ref).toLowerCase())) {
      throw new Error(`Destination interdite : "${ref}" (aucune suppression possible avec ce serveur).`);
    }
    await this.loadForbiddenIds();

    let current = await this.graph.get(`/me/mailFolders/${folderSegment(ref)}`, {
      query: { $select: 'id,displayName,parentFolderId' },
    });
    const target = current;
    for (let depth = 0; depth < 20 && current; depth++) {
      if (this.forbiddenIds.has(current.id)) {
        throw new Error(`Destination interdite : "${target.displayName}" est dans un dossier de suppression.`);
      }
      if (!current.parentFolderId || current.id === this.rootId || current.parentFolderId === this.rootId) break;
      current = await this.graph.get(`/me/mailFolders/${seg(current.parentFolderId)}`, {
        query: { $select: 'id,displayName,parentFolderId' },
      });
    }
    return target;
  }
}
