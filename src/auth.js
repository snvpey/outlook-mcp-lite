import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { PublicClientApplication, InteractionRequiredAuthError, LogLevel } from '@azure/msal-node';
import { SCOPES, LIMITS } from './config.js';

const LOGIN_HOST = 'login.microsoftonline.com';

/**
 * Cache de tokens chiffré via la librairie officielle Microsoft :
 * - Windows : fichier chiffré avec DPAPI (lié à ta session Windows)
 * - macOS   : Trousseau
 * - Linux   : libsecret
 * Si ce n'est pas disponible, on NE retombe PAS sur un fichier en clair :
 * les tokens restent en mémoire et il faut se reconnecter à chaque démarrage.
 */
async function createCachePlugin(config, log) {
  try {
    const ext = await import('@azure/msal-node-extensions');
    fs.mkdirSync(config.cacheDir, { recursive: true, mode: 0o700 });
    const persistence = await ext.PersistenceCreator.createPersistence({
      cachePath: path.join(config.cacheDir, 'token-cache.bin'),
      dataProtectionScope: ext.DataProtectionScope.CurrentUser,
      serviceName: 'outlook-mcp-lite',
      accountName: 'msal-token-cache',
      usePlaintextFileOnLinux: false,
    });
    return new ext.PersistenceCachePlugin(persistence);
  } catch (err) {
    if (config.requirePersistence) {
      throw new Error(`Stockage chiffré des tokens indisponible : ${err.message}`);
    }
    log(`Stockage chiffré des tokens indisponible (${err.message}). `
      + 'Tokens gardés en mémoire uniquement : reconnexion à chaque démarrage.');
    return null;
  }
}

/**
 * Ouvre le navigateur sans passer par un shell (pas d'injection possible
 * via l'URL) et uniquement vers la page de login Microsoft.
 */
export function openBrowser(url) {
  const parsed = new URL(url);
  if (parsed.protocol !== 'https:' || parsed.hostname !== LOGIN_HOST) {
    throw new Error(`URL de connexion inattendue : ${parsed.origin}`);
  }
  let cmd;
  let args;
  if (process.platform === 'win32') {
    cmd = 'rundll32';
    args = ['url.dll,FileProtocolHandler', url];
  } else if (process.platform === 'darwin') {
    cmd = 'open';
    args = [url];
  } else {
    cmd = 'xdg-open';
    args = [url];
  }
  const child = spawn(cmd, args, { stdio: 'ignore', detached: true, windowsHide: true });
  child.on('error', () => { /* l'URL est déjà affichée dans les logs */ });
  child.unref();
}

const SUCCESS_PAGE = '<html><body style="font-family:sans-serif;text-align:center;padding:40px">'
  + '<h2>Connexion réussie</h2><p>Tu peux fermer cette fenêtre.</p></body></html>';
const ERROR_PAGE = '<html><body style="font-family:sans-serif;text-align:center;padding:40px">'
  + '<h2>Échec de la connexion</h2><p>Ferme cette fenêtre et réessaie.</p></body></html>';

export class AuthManager {
  constructor(config, log = () => {}) {
    this.config = config;
    this.log = log;
    this.pca = null;
    this.pendingLogin = null;
  }

  async init() {
    if (this.pca) return;
    const cachePlugin = await createCachePlugin(this.config, this.log);
    this.pca = new PublicClientApplication({
      auth: {
        clientId: this.config.clientId,
        authority: this.config.authority,
      },
      cache: cachePlugin ? { cachePlugin } : undefined,
      system: {
        loggerOptions: {
          logLevel: LogLevel.Warning,
          piiLoggingEnabled: false,
          loggerCallback: (_level, message) => this.log(`[msal] ${message}`),
        },
      },
    });
  }

  async getAccount() {
    const accounts = await this.pca.getTokenCache().getAllAccounts();
    return accounts.find((a) => a.tenantId === this.config.tenantId) || null;
  }

  /** Renvoie un access token valide, en se reconnectant si besoin. */
  async getToken() {
    await this.init();
    const account = await this.getAccount();
    if (account) {
      try {
        const result = await this.pca.acquireTokenSilent({ account, scopes: SCOPES });
        return result.accessToken;
      } catch (err) {
        if (!(err instanceof InteractionRequiredAuthError)) throw err;
      }
    }
    return this.loginInteractive();
  }

  /** Une seule fenêtre de connexion à la fois, même si plusieurs outils sont appelés. */
  async loginInteractive() {
    if (!this.pendingLogin) {
      this.pendingLogin = this.runInteractive().finally(() => {
        this.pendingLogin = null;
      });
    }
    return this.pendingLogin;
  }

  async runInteractive() {
    // PKCE, state et serveur de retour sur 127.0.0.1 sont gérés par MSAL.
    const login = this.pca.acquireTokenInteractive({
      scopes: SCOPES,
      prompt: 'select_account',
      successTemplate: SUCCESS_PAGE,
      errorTemplate: ERROR_PAGE,
      openBrowser: async (url) => {
        this.log(`Connexion requise. Si le navigateur ne s'ouvre pas, va sur : ${url}`);
        openBrowser(url);
      },
    });

    let timer;
    const timeout = new Promise((_, reject) => {
      timer = setTimeout(
        () => reject(new Error('Connexion non terminée après 5 minutes, réessaie.')),
        LIMITS.loginTimeoutMs,
      );
      timer.unref();
    });

    try {
      const result = await Promise.race([login, timeout]);
      if (result.tenantId !== this.config.tenantId) {
        await this.pca.getTokenCache().removeAccount(result.account);
        throw new Error('Le compte choisi n\'appartient pas au tenant configuré.');
      }
      this.log(`Connecté : ${result.account?.username || 'compte inconnu'}`);
      return result.accessToken;
    } finally {
      clearTimeout(timer);
    }
  }

  async logout() {
    await this.init();
    const cache = this.pca.getTokenCache();
    for (const account of await cache.getAllAccounts()) {
      await cache.removeAccount(account);
    }
  }
}
