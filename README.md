# Hub

## Lancer le projet

1. **Installer Bun** : ouvrir PowerShell et coller :

   ```powershell
   powershell -c "irm bun.sh/install.ps1 | iex"
   ```

   Puis fermer et rouvrir PowerShell.

2. **Extraire le zip** du projet dans un dossier, par exemple `C:\Hub`.

3. **Lancer l'application** : dans PowerShell, coller (en remplaçant le chemin par celui du dossier extrait) :

   ```powershell
   cd C:\Hub; bun install; bun dev
   ```

4. **Ouvrir** [http://localhost:3300](http://localhost:3300) dans le navigateur.

Pour arrêter : `Ctrl + C` dans PowerShell. Les fois suivantes, seules les étapes 3 et 4 sont nécessaires.

Le serveur de jetons Microsoft (`localhost:3301`) démarre avec `bun dev` : la première fois, ouvrir
**Paramètres › Power BI** et se connecter avec le code affiché.

## Exports

- `bun run export:html` : toute l'application dans un seul fichier, `dist/hub.html`.
- `bun run export:server` : le serveur de jetons en un seul exécutable, `dist/hub-token-server`.

## Dépendances Microsoft

L'application appelle les API Microsoft depuis le navigateur. Les jetons viennent du serveur de jetons
local (`scripts/token-server.ts`) : une connexion par code d'appareil avec le client public Microsoft
Office, sans inscription d'application ni consentement administrateur.

| API | Accès | Utilisation |
| --- | --- | --- |
| Power BI REST API (`api.powerbi.com`) | Lecture des rapports et modèles ; requêtes DAX (`executeQueries`) ; actualisation lancée par l'utilisateur ; export `.pbix` | Liste et structure des rapports et modèles, chiffres des assistants |
| Power BI JavaScript SDK (`powerbi-client`) | Intégration des rapports (lecture et édition) | Rapports affichés, filtres, segments, pages, signets, export des données d'un visuel |
| Power BI Report Authoring (`powerbi-report-authoring`) | Édition du rapport intégré | Mode éditeur du Builder (rapports dont la définition n'est pas lisible), champs des visuels |
| Fabric REST API (`api.fabric.microsoft.com`) | Lecture et écriture des définitions (PBIR, TMDL) | Builder (pages, visuels, mesures de rapport), mesures et colonnes des modèles, nouveaux rapports |
| Microsoft Graph (`graph.microsoft.com`) | Fichiers OneDrive / SharePoint | Fichiers de contexte, images, options de l'application |

Bonnes pratiques de performance : [docs/performance.md](docs/performance.md).

## Crédits

Icône des contextes : [Flaticon](https://www.flaticon.com/free-icon/document_8243060).
