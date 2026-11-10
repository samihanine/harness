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

Le serveur local (`localhost:3301`) démarre avec `bun dev` : la première fois, ouvrir
**Paramètres › Power BI** et se connecter avec le code affiché. Il garde la session Microsoft et la clé
Confluence dans le trousseau du système, appelle les API Microsoft pour l'application et dessine les
PowerPoint (LibreOffice WebAssembly, installé par `bun install` ; plus rapide à démarrer si Node.js est installé).
Sécurité : [docs/security.md](docs/security.md).

Derrière un proxy d'entreprise qui inspecte le TLS : le serveur utilise les certificats du système
(rien à régler). Ne pas désactiver la vérification TLS.

## Vérifier

- `bun run typecheck` : les types et les frontières entre modules ([docs/structure.md](docs/structure.md)).
- `bun test` : conversions markdown ⇄ Confluence, présentations PowerPoint, tables Excel.
- `bun run screens` puis `bun run screens --check` (avec `bun dev` lancé) : ce que montre chaque écran, avant / après un changement.

## Exports

- `bun run export:html` : toute l'application dans un seul fichier, `dist/hub.html`, servi par le serveur local :
  ouvrir `http://localhost:3301` (ouvert comme fichier `file://`, il est refusé par le serveur, sauf `HUB_ALLOW_FILE=1`).
- `bun run export:server` : le serveur local en un seul exécutable, `dist/hub-token-server` (sans le moteur
  d'aperçu PowerPoint : l'application affiche alors son aperçu simplifié).

## Dépendances Microsoft

L'application appelle les API Microsoft à travers le serveur local (`scripts/token-server.ts`), qui ajoute
les jetons : une connexion par code d'appareil avec le client public Microsoft Office, sans inscription
d'application ni consentement administrateur. Seul le jeton Power BI arrive dans la page (en mémoire),
pour l'intégration des rapports. L'API d'IA est appelée par la page uniquement.

| API | Accès | Utilisation |
| --- | --- | --- |
| Power BI REST API (`api.powerbi.com`) | Lecture des rapports et modèles ; requêtes DAX (`executeQueries`) ; actualisation lancée par l'utilisateur ; export `.pbix` | Liste et structure des rapports et modèles, chiffres des assistants |
| Power BI JavaScript SDK (`powerbi-client`) | Intégration des rapports (lecture et édition) | Rapports affichés, filtres, segments, pages, signets, export des données d'un visuel |
| Power BI Report Authoring (`powerbi-report-authoring`) | Édition du rapport intégré | Mode éditeur du Builder (rapports dont la définition n'est pas lisible), champs des visuels |
| Fabric REST API (`api.fabric.microsoft.com`) | Lecture et écriture des définitions (PBIR, TMDL) | Builder (pages, visuels, mesures de rapport), mesures et colonnes des modèles, nouveaux rapports |
| Microsoft Graph (`graph.microsoft.com`) | Fichiers OneDrive / SharePoint ; API workbook d'Excel | Tables Excel (lignes, colonnes, mise en forme), présentations PowerPoint enregistrées sur SharePoint, options de l'application |
| Confluence REST API (via le serveur de jetons pour Confluence Cloud) | Lecture et écriture des pages, pièces jointes | Pages en markdown (diagrammes joints en images), contextes de l'IA |

Bonnes pratiques de performance : [docs/performance.md](docs/performance.md).

## Crédits

