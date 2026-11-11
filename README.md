# outlook-mcp-lite

Serveur MCP minimal pour connecter GitHub Copilot (ou un autre client MCP) à **ta propre boîte Outlook**, via Microsoft Graph.

Inspiré de [XenoXilus/outlook-mcp](https://github.com/XenoXilus/outlook-mcp), réécrit avec beaucoup moins de fonctionnalités et des garde-fous plus stricts.

## Ce que le serveur peut faire

| Outil | Action |
|---|---|
| `outlook_list_folders` | Lister les dossiers et sous-dossiers |
| `outlook_list_messages` | Lister les mails d'un dossier (résumé) |
| `outlook_search_messages` | Rechercher dans la boîte (KQL) |
| `outlook_get_message` | Lire un mail complet + liste des pièces jointes (nom, type, taille) |
| `outlook_list_categories` | Lister les catégories |
| `outlook_create_draft` | Créer un brouillon |
| `outlook_create_reply_draft` | Préparer une réponse (ou réponse à tous) en brouillon |
| `outlook_update_draft` | Modifier un brouillon (refuse si ce n'est pas un brouillon) |
| `outlook_create_folder` | Créer un dossier |
| `outlook_move_messages` | Déplacer des mails (max 25 par appel) |
| `outlook_create_category` | Créer une catégorie de couleur |
| `outlook_set_categories` | Ajouter / retirer des catégories sur des mails |
| `outlook_mark_read` | Marquer lu / non lu |
| `outlook_set_flag` | Drapeau de suivi |

## Ce qu'il ne peut PAS faire

- Envoyer, transférer ou répondre directement à un mail (seulement des brouillons)
- Supprimer un mail, un dossier ou une catégorie
- Déplacer un mail dans "Éléments supprimés" (ni dans un de ses sous-dossiers), la Boîte d'envoi ou les dossiers "Recoverable Items"
- Lire le contenu des pièces jointes
- Accéder à une autre boîte que la tienne, à OneDrive ou à SharePoint

## Modèle de sécurité

**Permissions minimales.** Le serveur demande seulement `User.Read`, `Mail.ReadWrite` et `MailboxSettings.ReadWrite` (plus `offline_access`, ajouté par MSAL). Pas de `Mail.Send` : même si le token était volé, il ne permettrait pas d'envoyer un mail.

**Liste blanche des appels Graph** (`src/graph.js`). Chaque requête est vérifiée avant d'être envoyée : méthode et chemin doivent correspondre à une route autorisée. `DELETE`, `/send`, `/forward`, `/reply`, `/permanentDelete`, `/users/...` sont refusés dans le code, même si un bug ailleurs essayait de les appeler. Les `PATCH` ne peuvent modifier que certains champs (lu, drapeau, catégories, et le contenu des brouillons).

**Protection contre la suppression déguisée** (`src/folders.js`). `Mail.ReadWrite` permet techniquement de "supprimer" un mail en le déplaçant dans la corbeille. Avant chaque déplacement, le serveur remonte la chaîne des dossiers parents de la destination et refuse si elle est dans un dossier de suppression.

**Tokens chiffrés, sans repli en clair** (`src/auth.js`). L'authentification utilise MSAL, la librairie officielle Microsoft (PKCE, `state`, serveur de retour sur `127.0.0.1`). Le cache de tokens est chiffré avec DPAPI sous Windows (Trousseau sous macOS, libsecret sous Linux) via `@azure/msal-node-extensions`. Si ce stockage n'est pas disponible, les tokens restent **en mémoire** et il faut se reconnecter à chaque démarrage. Il n'y a jamais de fichier de token avec une clé devinable, contrairement au repo d'origine. Pour refuser complètement de démarrer dans ce cas : `OUTLOOK_MCP_REQUIRE_PERSISTENCE=1`.

**Tenant verrouillé.** `AZURE_TENANT_ID` doit être le GUID de ton tenant (`common` est refusé) et le compte connecté doit appartenir à ce tenant.

**Contenu des mails marqué comme non fiable.** Le corps des mails est renvoyé entre deux marqueurs, avec une consigne demandant à l'IA de ne pas suivre les instructions qu'il contient. Les faux marqueurs présents dans un mail sont neutralisés. Ça réduit le risque de prompt injection sans le supprimer complètement (voir "Limites").

**Pas de parsing de fichiers.** Pas de xlsx, officeparser, DOMPurify ni Chrome. Graph convertit lui-même le HTML en texte. Quatre dépendances directes, toutes officielles (Microsoft, MCP, zod), versions figées, `npm audit` à 0 au moment de l'écriture.

**Validation des entrées.** Tous les paramètres passent par des schémas zod : format des ids, adresses email valides, tailles maximales, 25 mails max par opération groupée, 50 résultats max par liste.

**Logs.** Tout part sur stderr (stdout est réservé au protocole MCP). Ni token ni contenu de mail n'est loggé.

## Limites connues

- Les brouillons peuvent être influencés par un mail piégé (ex : un brouillon qui recopie des infos sensibles vers une adresse externe). Comme rien n'est envoyé automatiquement, **relis toujours un brouillon et ses destinataires avant de l'envoyer**.
- Le blocage des suppressions est dans le code du serveur, pas dans les permissions Graph (`Mail.ReadWrite` autorise `DELETE` côté API). Si tu modifies le code, garde la liste blanche.
- Le contenu des mails part vers le modèle utilisé par Copilot. Vérifie que c'est conforme à la politique de ton entreprise sur ce type de données.

## Installation

### 1. Prérequis

- Node.js 24 LTS recommandé (minimum 22.21, nécessaire pour `NODE_USE_ENV_PROXY`)
- Une "App Registration" dans Entra ID (Azure AD). En entreprise, il faut en général passer par la DSI.

### 2. App Registration (Entra ID)

1. Portail Azure, **Entra ID**, **Inscriptions d'applications**, **Nouvelle inscription**
2. Type de compte : **Comptes de cet annuaire organisationnel uniquement**
3. **Authentification**, **Ajouter une plateforme**, **Applications mobiles et de bureau**, URI de redirection : `http://localhost`
4. **Autorisations d'API**, **Microsoft Graph**, **Autorisations déléguées** : `User.Read`, `Mail.ReadWrite`, `MailboxSettings.ReadWrite`
5. Si le tenant l'exige : **Accorder le consentement administrateur** (à faire par un admin)
6. Note l'**ID d'application (client)** et l'**ID de l'annuaire (locataire)**

Aucun secret client n'est nécessaire.

### 3. Installer les dépendances (derrière Zscaler)

Zscaler déchiffre le trafic HTTPS avec son propre certificat. Node et npm ne lisent pas le magasin de certificats Windows, il faut leur donner ce certificat :

1. `certmgr.msc`, **Autorités de certification racines de confiance**, **Certificats**, clic droit sur le certificat Zscaler, **Toutes les tâches**, **Exporter**, format **Base 64 (.cer)**. Exemple : `C:\certs\zscaler-root.cer`
2. Dans PowerShell :

```powershell
npm config set cafile C:\certs\zscaler-root.cer
# Seulement si ton poste utilise un proxy explicite (pas nécessaire en mode tunnel Zscaler) :
npm config set https-proxy http://proxy.entreprise:port

cd C:\chemin\outlook-mcp-lite
npm ci
npm test
```

Si `npm ci` échoue sur `keytar` (module natif téléchargé depuis GitHub, utilisé seulement sur macOS/Linux mais requis par la librairie), lance `npm ci --ignore-scripts`. Le serveur fonctionnera alors en mode "tokens en mémoire" : connexion à chaque démarrage de VS Code, aucun token écrit sur le disque.

Ne mets jamais `NODE_TLS_REJECT_UNAUTHORIZED=0` ni `npm config set strict-ssl false` : ça désactive toute vérification TLS, y compris vers Microsoft.

### 4. Première connexion

```powershell
$env:AZURE_CLIENT_ID = "<client id>"
$env:AZURE_TENANT_ID = "<tenant id>"
$env:NODE_EXTRA_CA_CERTS = "C:\certs\zscaler-root.cer"
$env:NODE_USE_ENV_PROXY = "1"            # seulement avec un proxy explicite
$env:HTTPS_PROXY = "http://proxy:port"   # seulement avec un proxy explicite
npm run login
```

Le navigateur s'ouvre sur la page de connexion Microsoft. Une fois connecté, le script affiche ton nom. Pour te déconnecter : `npm run logout`.

### 5. Configuration dans VS Code (GitHub Copilot)

Fichier `.vscode/mcp.json` (ou commande **MCP: Open User Configuration**) :

```json
{
  "servers": {
    "outlook": {
      "type": "stdio",
      "command": "node",
      "args": ["C:\\chemin\\outlook-mcp-lite\\src\\index.js"],
      "env": {
        "AZURE_CLIENT_ID": "<client id>",
        "AZURE_TENANT_ID": "<tenant id>",
        "NODE_EXTRA_CA_CERTS": "C:\\certs\\zscaler-root.cer",
        "NODE_USE_ENV_PROXY": "1",
        "HTTPS_PROXY": "http://proxy:port"
      }
    }
  }
}
```

Retire `NODE_USE_ENV_PROXY` et `HTTPS_PROXY` si Zscaler fonctionne en mode tunnel.

Conseils côté Copilot :
- Laisse la confirmation des appels d'outils activée, surtout pour les outils d'écriture. N'active pas l'auto-approve.
- Tu peux désactiver des outils un par un dans la liste des outils du chat si tu veux un mode lecture seule.

## Variables d'environnement

| Variable | Obligatoire | Rôle |
|---|---|---|
| `AZURE_CLIENT_ID` | oui | ID de l'App Registration |
| `AZURE_TENANT_ID` | oui | GUID du tenant |
| `NODE_EXTRA_CA_CERTS` | derrière Zscaler | Certificat racine Zscaler |
| `NODE_USE_ENV_PROXY` + `HTTPS_PROXY` | proxy explicite | Fait passer `fetch` par le proxy |
| `OUTLOOK_MCP_CACHE_DIR` | non | Dossier du cache chiffré (défaut : `~/.outlook-mcp-lite`) |
| `OUTLOOK_MCP_REQUIRE_PERSISTENCE` | non | `1` = refuse de démarrer sans stockage chiffré |

## Structure

```
src/
  index.js        démarrage stdio
  server.js       création du serveur MCP et des outils
  config.js       configuration, scopes, limites
  auth.js         MSAL, cache chiffré, ouverture du navigateur
  graph.js        client Graph + liste blanche des routes
  folders.js      protection des dossiers de suppression
  format.js       mise en forme, marquage du contenu non fiable
  tools/
    schemas.js    schémas zod partagés
    read.js       outils de lecture
    write.js      outils d'écriture (brouillons, tri)
scripts/login.js  connexion / déconnexion en ligne de commande
test/             tests (node --test)
```
