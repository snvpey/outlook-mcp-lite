#!/usr/bin/env node
// Connexion initiale (ou déconnexion) depuis un terminal, hors de VS Code.
//   npm run login    -> ouvre le navigateur, enregistre le token chiffré
//   npm run logout   -> supprime les tokens du cache
import { loadConfig } from '../src/config.js';
import { AuthManager } from '../src/auth.js';
import { GraphClient } from '../src/graph.js';

const log = (...args) => console.error(...args);

async function main() {
  const config = loadConfig();
  const auth = new AuthManager(config, log);

  if (process.argv.includes('--logout')) {
    await auth.logout();
    log('Déconnecté : tokens supprimés du cache.');
    return;
  }

  const graph = new GraphClient(auth);
  const me = await graph.get('/me', { query: { $select: 'displayName,userPrincipalName' } });
  log(`OK, connecté en tant que ${me.displayName} (${me.userPrincipalName}).`);
}

main().catch((err) => {
  log(`Échec : ${err.message}`);
  process.exit(1);
});
