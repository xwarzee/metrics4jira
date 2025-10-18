# Installation Guide - Jira JQL Explorer pour Excel 2016

Guide d'installation détaillé pour intégrer les macros VBA dans Excel 2016.

## Prérequis

- Microsoft Excel 2016 ou supérieur (Windows ou Mac)
- Accès à un compte Jira avec un token API
- Connexion Internet

## Étapes d'Installation

### 1. Activer l'onglet Développeur dans Excel

#### Sur Windows :
1. Ouvrez Excel
2. Cliquez sur **Fichier** > **Options**
3. Sélectionnez **Personnaliser le ruban**
4. Dans la colonne de droite, cochez **Développeur**
5. Cliquez sur **OK**

#### Sur Mac :
1. Ouvrez Excel
2. Cliquez sur **Excel** (menu) > **Préférences**
3. Sélectionnez **Ruban et barre d'outils**
4. Dans l'onglet **Ruban**, cochez **Développeur**
5. Fermez les préférences

### 2. Créer un nouveau classeur Excel

1. Créez un nouveau classeur Excel vierge
2. Enregistrez-le sous le nom **JiraExplorer.xlsm** (format Macro Excel)
   - **Fichier** > **Enregistrer sous**
   - Type de fichier : **Classeur Excel avec macros (*.xlsm)**

### 3. Importer les modules VBA

1. Appuyez sur **Alt + F11** (Windows) ou **Fn + Option + F11** (Mac) pour ouvrir l'éditeur VBA

2. Dans l'éditeur VBA, pour chaque fichier .bas :
   - Cliquez sur **Fichier** > **Importer un fichier**
   - Sélectionnez le fichier :
     - `JiraConfig.bas`
     - `JiraApiClient.bas`
     - `JiraExplorer.bas`
   - Répétez pour les 3 fichiers

3. Vous devriez voir 3 modules dans le volet de gauche :
   - JiraConfig
   - JiraApiClient
   - JiraExplorer

### 4. Ajouter les références nécessaires

1. Dans l'éditeur VBA, cliquez sur **Outils** > **Références**

2. Cochez les bibliothèques suivantes :
   - ✅ **Microsoft XML, v6.0** (pour les requêtes HTTP)
   - ✅ **Microsoft Scripting Runtime** (pour Dictionary)
   - ✅ **Microsoft Script Control 1.0** (pour parser JSON)

3. Cliquez sur **OK**

**Note pour Mac** : Si "Microsoft Script Control" n'est pas disponible sur Mac, vous devrez utiliser une approche alternative pour le parsing JSON (voir section Dépannage).

### 5. Créer l'interface avec boutons

1. Fermez l'éditeur VBA et revenez à Excel

2. Dans l'onglet **Développeur**, cliquez sur **Insérer** > **Bouton (Contrôle de formulaire)**

3. Créez 4 boutons sur la première feuille et assignez les macros :

   **Bouton 1 : "Initialize Workbook"**
   - Texte du bouton : `Initialize Workbook`
   - Macro assignée : `JiraExplorer.InitializeWorkbook`

   **Bouton 2 : "Configure Connection"**
   - Texte du bouton : `Configure Connection`
   - Macro assignée : `JiraExplorer.ConfigureJiraConnection`

   **Bouton 3 : "Test Connection"**
   - Texte du bouton : `Test Connection`
   - Macro assignée : `JiraExplorer.TestJiraConnection`

   **Bouton 4 : "Search Jira Issues"**
   - Texte du bouton : `Search Jira Issues`
   - Macro assignée : `JiraExplorer.SearchJiraIssues`

4. Disposez les boutons de manière lisible

### 6. Initialiser le classeur

1. Cliquez sur le bouton **"Initialize Workbook"**
2. Trois feuilles seront créées automatiquement :
   - **Config** : Configuration de la connexion Jira
   - **Issues** : Liste des issues trouvées
   - **FieldExplorer** : Détails d'une issue sélectionnée

### 7. Configurer la connexion Jira

1. Allez dans la feuille **Config**

2. Remplissez les informations :
   - **Jira URL** : L'URL de votre instance Jira
     - Exemple : `https://votre-domaine.atlassian.net`
   - **Username (Email)** : Votre email Jira
   - **API Token** : Votre token API (voir ci-dessous)
   - **Max Results** : Nombre maximum de résultats (1-1000)
   - **API Version** : Sélectionnez dans la liste déroulante
     - `Jira Server 9.12.24` pour Jira on-premise
     - `Jira Cloud (Current)` pour Jira Cloud

### 8. Générer un token API Jira

1. Allez sur : https://id.atlassian.com/manage-profile/security/api-tokens
2. Cliquez sur **"Create API token"**
3. Donnez un nom au token (ex: "Excel JQL Explorer")
4. Copiez le token généré
5. Collez-le dans la cellule **B4** de la feuille Config

**Important** : Le token n'est affiché qu'une seule fois. Conservez-le en lieu sûr.

### 9. Tester la connexion

1. Cliquez sur le bouton **"Test Connection"**
2. Si la connexion réussit, un message de confirmation s'affiche
3. Si la connexion échoue, vérifiez vos paramètres

## Utilisation

### Rechercher des issues Jira

1. Cliquez sur **"Search Jira Issues"**
2. Entrez une requête JQL dans la boîte de dialogue

Exemples de requêtes JQL :
```
project = MYPROJECT
assignee = currentUser()
status = Open AND priority = High
created >= -7d
status IN (Open, "In Progress") AND assignee = currentUser() ORDER BY created DESC
```

3. Les résultats s'affichent dans la feuille **Issues**

### Explorer les détails d'une issue

1. Dans la feuille **Issues**, cliquez sur une ligne d'issue
2. Exécutez la macro `JiraExplorer.ShowIssueDetails` avec le numéro de ligne
3. Les détails s'affichent dans la feuille **FieldExplorer**

**Astuce** : Vous pouvez créer un bouton qui appelle cette macro ou l'assigner à un raccourci clavier.

## Configuration des Paramètres de Sécurité

### Windows

1. **Fichier** > **Options** > **Centre de gestion de la confidentialité**
2. Cliquez sur **Paramètres du Centre de gestion de la confidentialité**
3. Sélectionnez **Paramètres des macros**
4. Choisissez **Activer toutes les macros** (ou **Désactiver toutes les macros avec notification**)
5. Cochez **Faire confiance à l'accès au modèle objet du projet VBA**
6. Cliquez sur **OK**

### Mac

1. **Excel** > **Préférences** > **Sécurité et confidentialité**
2. Dans **Sécurité des macros**, choisissez :
   - **Activer toutes les macros** (recommandé pour le développement)
   - Ou **Désactiver toutes les macros avec notification**

## Dépannage

### Erreur "Référence manquante"

Si vous obtenez une erreur de référence manquante :

1. Ouvrez l'éditeur VBA (**Alt + F11**)
2. **Outils** > **Références**
3. Décochez les références marquées comme **MANQUANT**
4. Recherchez et cochez la version correcte de la bibliothèque

### Erreur "Microsoft Script Control non disponible" (Mac)

Sur Mac, ScriptControl n'est pas disponible. Solutions :

**Option 1** : Utiliser une bibliothèque JSON VBA tierce
- Téléchargez VBA-JSON : https://github.com/VBA-tools/VBA-JSON
- Importez le module `JsonConverter.bas`
- Modifiez `JiraApiClient` pour utiliser `JsonConverter.ParseJson`

**Option 2** : Travailler uniquement sur Windows
- Ouvrez le fichier sur un PC Windows où ScriptControl est disponible

### Erreur "Impossible de se connecter à Jira"

Vérifiez :
- L'URL Jira est correcte (doit commencer par `https://`)
- Le token API est valide
- Votre réseau autorise les connexions HTTPS sortantes
- Votre pare-feu/proxy ne bloque pas les connexions

### Erreur "Invalid request payload" (API v3)

Si vous utilisez Jira Cloud et obtenez cette erreur :
- Assurez-vous d'avoir sélectionné **"Jira Cloud (Current)"** dans API Version
- L'API v3 utilise GET avec query parameters
- L'API v2 utilise POST avec JSON body

### Performance lente

Si les recherches sont lentes :
- Réduisez le **Max Results** dans la config
- Affinez vos requêtes JQL pour limiter les résultats
- Vérifiez votre connexion Internet

## Différences API v2 vs v3

### Jira Server 9.12.24 (API v2)
- Endpoint : `/rest/api/2/search`
- Méthode : **POST** avec JSON body
- Fields : `["*all"]`
- Expand : `"names,schema"`

### Jira Cloud (API v3)
- Endpoint : `/rest/api/3/search/jql`
- Méthode : **GET** avec query parameters
- Fields : `*navigable`
- Expand : Non utilisé

## Sauvegarde et Distribution

### Sauvegarder votre configuration

La configuration est stockée dans la feuille "Config" du classeur. Pour sauvegarder :
1. Enregistrez le classeur (**Ctrl + S** ou **Cmd + S**)
2. La configuration sera conservée

**Attention** : Le token API est stocké en clair dans le classeur. Ne partagez pas le fichier avec le token.

### Distribuer à d'autres utilisateurs

Pour partager avec des collègues :
1. Supprimez le contenu de la feuille **Config** (sauf les en-têtes)
2. Enregistrez une copie du classeur
3. Distribuez cette copie
4. Chaque utilisateur devra configurer ses propres identifiants

## Support et Ressources

- **Documentation Jira API** : https://developer.atlassian.com/cloud/jira/platform/rest/v3/
- **Guide JQL** : https://support.atlassian.com/jira-software-cloud/docs/what-is-advanced-searching-in-jira-cloud/
- **Génération de tokens API** : https://id.atlassian.com/manage-profile/security/api-tokens

## Notes de Version

### Version 1.0
- Support Jira Server 9.12.24 (API v2)
- Support Jira Cloud (API v3)
- Recherche JQL avec pagination
- Explorateur de champs
- Configuration dans Excel
- Compatible Excel 2016+
