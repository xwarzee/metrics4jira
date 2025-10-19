# Jira JQL Explorer - Version Excel 2016

Application Excel avec macros VBA pour explorer les issues Jira via des requêtes JQL directement dans Excel.

## Vue d'ensemble

Cette version Excel permet d'interroger Jira et d'afficher les résultats directement dans des feuilles Excel, sans avoir besoin d'installer de logiciel supplémentaire. Parfait pour les utilisateurs qui travaillent principalement dans Excel et souhaitent intégrer des données Jira dans leurs analyses.

## Fonctionnalités

- ✅ **Requêtes JQL personnalisées** : Recherchez des issues avec n'importe quelle requête JQL
- ✅ **Interface Excel native** : Boutons et feuilles Excel pour une expérience familière
- ✅ **Support dual API** : Compatible avec Jira Server 9.12.24 (API v2) et Jira Cloud (API v3)
- ✅ **Explorateur de champs** : Visualisez tous les champs d'une issue
- ✅ **Configuration intégrée** : Stockez votre configuration dans le classeur
- ✅ **Test de connexion** : Vérifiez votre configuration avant de lancer des recherches
- ✅ **Métadonnées des champs** : Noms de champs lisibles automatiquement chargés
- ✅ **Support Proxy** : Configuration proxy HTTP/HTTPS avec authentification
- ✅ **Aucune installation requise** : Fonctionne avec Excel 2016+ sur Windows uniquement

## Prérequis

- **Microsoft Excel 2016 ou supérieur sur Windows** (⚠️ Version Windows uniquement)
- Un compte Jira avec un token API
- Connexion Internet

### ⚠️ Important : Compatibilité macOS

**La version Excel VBA ne fonctionne PAS sur macOS** en raison de limitations techniques :
- Les objets MSXML (requêtes HTTP) ne sont pas disponibles sur Excel Mac
- ScriptControl (parsing JSON) n'existe pas sur Mac

**Alternatives pour macOS** :
- ✅ Utilisez la **version Java** (dans le dossier `/java`)
- ✅ Utilisez la **version Python** (dans le dossier `/python`)
- ⚙️ Utilisez Windows via Parallels, VMware ou Boot Camp

## Installation Rapide

1. **Créer un nouveau classeur Excel** et l'enregistrer en format `.xlsm` (avec macros)
2. **Importer les 3 modules VBA** dans l'éditeur VBA (Alt + F11) :
   - `JiraConfig.bas`
   - `JiraApiClient.bas`
   - `JiraExplorer.bas`
3. **Ajouter les références** (Outils > Références) :
   - Microsoft XML, v6.0
   - Microsoft Scripting Runtime
   - Microsoft Script Control 1.0 (Windows uniquement)
4. **Créer 4 boutons** dans la feuille Excel et assigner les macros
5. **Initialiser le classeur** en cliquant sur "Initialize Workbook"

Pour des instructions détaillées, consultez [INSTALLATION.md](INSTALLATION.md).

## Structure des Fichiers

```
excel/
├── JiraConfig.bas          # Module de configuration
├── JiraApiClient.bas       # Client API REST Jira
├── JiraExplorer.bas        # Interface principale Excel
├── INSTALLATION.md         # Guide d'installation détaillé
└── README.md              # Ce fichier
```

## Utilisation

### 1. Configuration initiale

1. Cliquez sur **"Initialize Workbook"** pour créer les feuilles nécessaires
2. Allez dans la feuille **Config**
3. Remplissez :
   - **Jira URL** : `https://votre-domaine.atlassian.net`
   - **Username** : Votre email Jira
   - **API Token** : Généré depuis https://id.atlassian.com/manage-profile/security/api-tokens
   - **Max Results** : 50 (ou autre valeur entre 1-1000)
   - **API Version** : Choisissez votre version Jira

### 2. Configuration du Proxy (Optionnel)

Si vous êtes derrière un proxy d'entreprise :
1. **Use Proxy** : Sélectionnez "Yes"
2. **Proxy Server** : `proxy.company.com` (adresse de votre proxy)
3. **Proxy Port** : `8080` (ou le port utilisé par votre proxy)
4. **Proxy Username** : Votre identifiant proxy (si authentification requise)
5. **Proxy Password** : Votre mot de passe proxy (si authentification requise)

**Exemples de configuration proxy** :
```
Proxy sans authentification :
- Use Proxy: Yes
- Proxy Server: proxy.company.com
- Proxy Port: 8080
- Proxy Username: (vide)
- Proxy Password: (vide)

Proxy avec authentification :
- Use Proxy: Yes
- Proxy Server: proxy.company.com
- Proxy Port: 3128
- Proxy Username: john.doe
- Proxy Password: ••••••••
```

### 3. Test de connexion

Cliquez sur **"Test Connection"** pour vérifier que la connexion fonctionne (avec ou sans proxy).

### 4. Recherche d'issues

1. Cliquez sur **"Search Jira Issues"**
2. Entrez votre requête JQL :
   ```
   project = MYPROJECT AND status = Open
   ```
3. Les résultats s'affichent dans la feuille **Issues**

### 5. Explorer les détails

Pour voir tous les champs d'une issue :
1. Notez le numéro de ligne de l'issue dans la feuille **Issues**
2. Exécutez `JiraExplorer.ShowIssueDetails(numéro_ligne)` depuis VBA
3. Les détails s'affichent dans **FieldExplorer**

## Exemples de Requêtes JQL

```jql
// Issues d'un projet
project = MYPROJECT

// Issues assignées à moi
assignee = currentUser()

// Issues créées dans les 7 derniers jours
created >= -7d

// Issues ouvertes avec priorité haute
status = Open AND priority = High

// Issues dans plusieurs statuts
status IN (Open, "In Progress", "To Do")

// Combinaison avec tri
project = MYPROJECT AND status != Done AND assignee = currentUser() ORDER BY priority DESC
```

## Architecture VBA

### Module JiraConfig
Gère la configuration de connexion Jira :
- Enum `ApiVersion` : SERVER_9_12_24 et CLOUD_CURRENT
- Type `JiraConfiguration` : Structure de configuration
- Fonctions de chargement/sauvegarde depuis la feuille Config
- Génération du header Authorization (Base64)
- Gestion des endpoints API v2 vs v3

### Module JiraApiClient
Gère les appels API REST Jira :
- `TestConnection()` : Teste la connexion Jira
- `SearchIssues()` : Exécute une requête JQL
- `SearchIssuesCloud()` : Méthode GET pour API v3
- `SearchIssuesServer()` : Méthode POST pour API v2
- `GetFieldMetadata()` : Charge les métadonnées des champs
- Parsing JSON avec ScriptControl (Windows) ou JScript

### Module JiraExplorer
Interface utilisateur Excel :
- `InitializeWorkbook()` : Crée les feuilles nécessaires
- `ConfigureJiraConnection()` : Ouvre la feuille Config
- `TestJiraConnection()` : Teste la connexion
- `SearchJiraIssues()` : Lance une recherche JQL
- `ShowIssueDetails()` : Affiche les détails d'une issue
- Mise en forme automatique des feuilles

## API Jira REST

### Différences entre v2 (Server) et v3 (Cloud)

| Aspect | API v2 (Server 9.12.24) | API v3 (Cloud) |
|--------|-------------------------|----------------|
| Endpoint | `/rest/api/2/search` | `/rest/api/3/search/jql` |
| Méthode HTTP | POST | GET |
| Body/Params | JSON body | Query parameters |
| Fields | `["*all"]` | `*navigable` |
| Headers XSRF | `X-Atlassian-Token: no-check` | Non requis |

L'application gère automatiquement ces différences en fonction de la version API sélectionnée.

## Sécurité

⚠️ **Important** : Le token API est stocké en clair dans la feuille Config du classeur Excel.

**Recommandations** :
- Ne partagez jamais le classeur contenant votre token API
- Stockez le fichier dans un emplacement sécurisé
- Utilisez des permissions de fichier appropriées
- Pour distribuer le classeur, supprimez d'abord la configuration

**Alternative sécurisée** :
- Créez un fichier Config externe (non inclus dans le classeur)
- Modifiez le code pour charger la config depuis ce fichier externe

## Limitations

### Limitations Excel VBA
- **Parsing JSON** : Utilise ScriptControl (Windows uniquement) ou nécessite une bibliothèque tierce sur Mac
- **Requêtes asynchrones** : VBA est monothread, les recherches bloquent Excel pendant l'exécution
- **Taille des données** : Excel limite à ~1 million de lignes
- **Performance** : Plus lent que les versions Java ou Python pour de gros volumes

### Limitations API
- **Pagination** : Jira limite les résultats par page (max 1000)
- **Rate limiting** : Jira Cloud impose des limites de taux
- **Timeout** : Les requêtes très longues peuvent timeout

### Compatibilité
- **Windows** : ✅ Pleine compatibilité avec Excel 2016+
- **Mac** : ❌ Non compatible (MSXML et ScriptControl non disponibles)
  - **Alternative** : Utilisez les versions Java ou Python
- **Excel Online** : ❌ Non compatible (VBA non supporté)

## Dépannage

### Erreur "Référence manquante"
**Solution** : Ouvrez l'éditeur VBA, allez dans Outils > Références, et vérifiez que toutes les bibliothèques sont cochées et disponibles.

### Erreur "Invalid request payload" (API v3)
**Solution** : Assurez-vous d'avoir sélectionné "Jira Cloud (Current)" dans la configuration API Version.

### Erreur sur macOS
**Problème** : MSXML et ScriptControl ne sont pas disponibles sur Excel Mac.
**Solution** : La version Excel VBA ne fonctionne que sur Windows. Pour macOS, utilisez :
- Version Java (dans `/java`)
- Version Python (dans `/python`)

### Performance lente
**Solution** : Réduisez Max Results dans la configuration ou affinez vos requêtes JQL.

### Connexion échoue
**Solution** : Vérifiez l'URL, le token API, et que votre réseau autorise les connexions HTTPS sortantes.

## Avantages par Rapport aux Autres Versions

### Vs Version Java
- ✅ **Pas d'installation** : Excel déjà installé sur les postes Windows
- ✅ **Intégration Excel** : Données directement dans Excel pour analyses
- ✅ **Courbe d'apprentissage** : Interface familière
- ❌ **Windows uniquement** : Ne fonctionne pas sur macOS
- ❌ **Performance** : Plus lent pour de gros volumes
- ❌ **Fonctionnalités** : Moins de features avancées

### Vs Version Python
- ✅ **Pas d'environnement Python** : Fonctionne avec juste Excel (Windows)
- ✅ **Interface Excel** : Idéal pour les utilisateurs Excel
- ❌ **Windows uniquement** : Ne fonctionne pas sur macOS
- ❌ **Flexibilité** : Moins flexible que Python pour scripting
- ❌ **Automatisation** : Moins adapté pour l'automatisation

### Choix de la Version

**Utilisez Excel VBA si** :
- ✅ Vous êtes sur Windows
- ✅ Vous travaillez principalement dans Excel
- ✅ Vous voulez une solution sans installation

**Utilisez Java ou Python si** :
- ✅ Vous êtes sur macOS
- ✅ Vous avez besoin de performances
- ✅ Vous voulez automatiser des tâches

## Développement

### Ajouter de nouvelles fonctionnalités

1. **Ouvrir l'éditeur VBA** : Alt + F11 (Windows) ou Fn + Option + F11 (Mac)
2. **Modifier le module** concerné
3. **Tester** : F5 pour exécuter une macro
4. **Déboguer** : F8 pour pas-à-pas, Ctrl + G pour la fenêtre Exécution

### Bonnes pratiques VBA
- Utilisez `Option Explicit` en haut de chaque module
- Gérez les erreurs avec `On Error GoTo ErrorHandler`
- Libérez les objets avec `Set obj = Nothing`
- Commentez votre code
- Testez avec de petits datasets d'abord

### Architecture modulaire
Les 3 modules sont séparés pour faciliter la maintenance :
- **JiraConfig** : Configuration pure, aucune logique métier
- **JiraApiClient** : Logique API, indépendant de l'interface
- **JiraExplorer** : Interface Excel, utilise les autres modules

## Ressources

- [Documentation Jira REST API v3](https://developer.atlassian.com/cloud/jira/platform/rest/v3/)
- [Documentation Jira REST API v2](https://docs.atlassian.com/software/jira/docs/api/REST/9.12.0/)
- [Guide JQL](https://support.atlassian.com/jira-software-cloud/docs/what-is-advanced-searching-in-jira-cloud/)
- [Génération de tokens API](https://id.atlassian.com/manage-profile/security/api-tokens)
- [VBA-JSON pour Mac](https://github.com/VBA-tools/VBA-JSON)

## Contribution

Pour contribuer à ce projet :
1. Ajoutez des fonctionnalités dans les modules VBA
2. Testez sur Windows et Mac si possible
3. Documentez les changements
4. Partagez vos améliorations

## Licence

Ce projet est fourni à des fins éducatives et de démonstration.

## Auteur

Version Excel VBA du Jira JQL Explorer, compatible avec Excel 2016 et supérieur.

## Changelog

### Version 1.0 (2025)
- Support Jira Server 9.12.24 (API v2)
- Support Jira Cloud (API v3)
- Interface Excel avec 3 feuilles (Config, Issues, FieldExplorer)
- Configuration dans Excel
- Test de connexion
- Recherche JQL
- Explorateur de champs
- Métadonnées des champs
- Compatible Windows et Mac (avec adaptation)
