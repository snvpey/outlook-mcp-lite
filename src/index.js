#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { loadConfig } from './config.js';
import { AuthManager } from './auth.js';
import { createServer } from './server.js';

// stdout est réservé au protocole MCP : tout log part sur stderr.
const log = (...args) => console.error('[outlook-mcp-lite]', ...args);
console.log = log;
console.info = log;

process.on('unhandledRejection', (err) => log(`Rejet non géré : ${err?.message || err}`));

async function main() {
  let config;
  try {
    config = loadConfig();
  } catch (err) {
    log(`Configuration invalide : ${err.message}`);
    process.exit(1);
  }

  const auth = new AuthManager(config, log);
  // Prépare le cache de tokens dès le démarrage pour signaler tout de suite
  // si le stockage chiffré n'est pas disponible.
  await auth.init();

  const server = createServer({ auth, log });
  await server.connect(new StdioServerTransport());
  log('Serveur démarré (stdio).');
}

main().catch((err) => {
  log(`Erreur fatale : ${err.message}`);
  process.exit(1);
});
