# metrics4jira

# Jira JQL Explorer - Interactive Whiteboard

Application Java interactive pour explorer les issues Jira via des requêtes JQL et visualiser tous les champs de manière détaillée.

## Fonctionnalités

- **Requêtes JQL personnalisées** : Exécutez n'importe quelle requête JQL pour rechercher des issues Jira
- **Whiteboard interactif** : Interface graphique moderne avec JavaFX
- **Explorateur de champs** : Visualisez tous les champs de chaque issue avec leurs valeurs
- **Vue JSON brute** : Affichez la réponse JSON complète de l'API Jira
- **Métadonnées des champs** : Noms de champs lisibles grâce aux métadonnées Jira
- **Copie facile** : Copiez les valeurs des champs ou le JSON complet dans le presse-papiers
- **Compatible Jira 9.12.24** : Utilise l'API REST v3 de Jira

## Prérequis

- Java 17 ou supérieur
- Maven 3.6 ou supérieur
- Un compte Jira avec un token API

## Installation

1. Clonez ou téléchargez ce projet :
```bash
cd metrics4jira
```

2. Compilez le projet avec Maven :
```bash
mvn clean package
```

## Configuration

### Méthode 1 : Via l'interface utilisateur (Recommandé)

1. Lancez l'application
2. Cliquez sur le bouton "Configure"
3. Remplissez les informations :
   - **Jira URL** : L'URL de votre instance Jira (ex: https://votre-domaine.atlassian.net)
   - **Username** : Votre email Jira
   - **API Token** : Votre token API Jira
   - **Max Results** : Nombre maximum de résultats par requête (1-1000)

### Méthode 2 : Via le fichier de configuration

Éditez le fichier `src/main/resources/jira.properties` :

```properties
jira.url=https://votre-domaine.atlassian.net
jira.username=votre-email@example.com
jira.apitoken=votre_token_api_jira
jira.maxresults=50
```

### Générer un token API Jira

1. Allez sur https://id.atlassian.com/manage-profile/security/api-tokens
2. Cliquez sur "Create API token"
3. Donnez un nom au token (ex: "JQL Explorer")
4. Copiez le token généré

**Important** : Le token n'est affiché qu'une seule fois. Conservez-le en lieu sûr.

## Utilisation

### Lancer l'application

#### Avec Maven :
```bash
mvn javafx:run
```

#### Avec le JAR compilé :
```bash
java -jar target/jira-jql-explorer-1.0.0.jar
```

### Utilisation de l'interface

1. **Configuration** : Cliquez sur "Configure" pour configurer votre connexion Jira

2. **Requête JQL** : Entrez votre requête JQL dans le champ de texte. Exemples :
   ```
   project = MYPROJECT AND status = Open
   assignee = currentUser() ORDER BY created DESC
   status IN (Open, "In Progress") AND priority = High
   created >= -7d
   ```

3. **Recherche** : Cliquez sur "Search" ou appuyez sur Entrée

4. **Exploration** :
   - La liste des issues apparaît à gauche
   - Cliquez sur une issue pour voir ses détails
   - L'explorateur de champs affiche tous les champs avec leurs valeurs
   - Le panneau JSON brut montre la structure complète

5. **Copie** :
   - Sélectionnez un champ et cliquez sur "Copy Field Value"
   - Cliquez sur "Copy JSON" pour copier le JSON complet

## Structure du projet

```
metrics4jira/
├── pom.xml                                 # Configuration Maven
├── README.md                               # Ce fichier
└── src/
    └── main/
        ├── java/
        │   └── com/jira/explorer/
        │       ├── JiraExplorerApp.java    # Classe principale
        │       ├── model/
        │       │   ├── JiraConfig.java     # Configuration Jira
        │       │   └── JiraIssue.java      # Modèle d'issue
        │       ├── service/
        │       │   └── JiraApiClient.java  # Client API REST Jira
        │       └── ui/
        │           ├── MainViewController.java  # Contrôleur principal
        │           └── ConfigDialog.java        # Dialog de configuration
        └── resources/
            ├── jira.properties             # Configuration Jira
            ├── styles.css                  # Styles CSS
            └── simplelogger.properties     # Configuration logging
```

## Dépendances principales

- **JavaFX 21.0.1** : Interface graphique moderne
- **OkHttp 4.12.0** : Client HTTP pour l'API Jira
- **Gson 2.10.1** : Parsing et manipulation JSON
- **SLF4J 2.0.9** : Logging
- **ControlsFX 11.1.2** : Contrôles UI avancés

## API Jira REST v3

Cette application utilise l'API REST v3 de Jira, compatible avec Jira 9.12.24 :

- **Endpoint de recherche** : `/rest/api/3/search`
- **Endpoint des champs** : `/rest/api/3/field`
- **Authentification** : Basic Auth avec username et API token
- **Format** : JSON

### Exemples de requêtes JQL

```jql
# Issues d'un projet
project = MYPROJECT

# Issues assignées à moi
assignee = currentUser()

# Issues créées dans les 7 derniers jours
created >= -7d

# Issues ouvertes avec priorité haute
status = Open AND priority = High

# Issues dans plusieurs statuts
status IN (Open, "In Progress", "To Do")

# Combinaison de filtres
project = MYPROJECT AND status != Done AND assignee = currentUser() ORDER BY priority DESC
```

## Dépannage

### L'application ne se lance pas

- Vérifiez que Java 17 ou supérieur est installé : `java -version`
- Vérifiez que Maven est installé : `mvn -version`
- Recompilez le projet : `mvn clean package`

### Erreur de connexion à Jira

- Vérifiez que l'URL Jira est correcte (doit commencer par https://)
- Vérifiez que le token API est valide
- Vérifiez que votre compte a les permissions nécessaires
- Testez la connexion avec curl :
  ```bash
  curl -u votre-email@example.com:votre_token https://votre-domaine.atlassian.net/rest/api/3/myself
  ```

### Erreur JQL

- Vérifiez la syntaxe JQL dans la documentation Jira
- Testez la requête directement dans Jira (Filtres > Recherche avancée)
- Les noms de champs personnalisés doivent être entre crochets : `cf[10001]`

### Problème d'affichage JavaFX

- Assurez-vous que JavaFX est correctement installé
- Sur macOS, vous pourriez avoir besoin d'installer JavaFX séparément

## Développement

### Compiler le projet

```bash
mvn clean compile
```

### Exécuter les tests (si présents)

```bash
mvn test
```

### Créer un JAR exécutable

```bash
mvn clean package
```

Le JAR sera créé dans `target/jira-jql-explorer-1.0.0.jar`

## Licence

Ce projet est fourni à des fins éducatives et de démonstration.

## Auteur

Créé par Claude Code - Application experte pour l'API REST Jira 9.12.24

## Support

Pour toute question ou problème :
1. Vérifiez la documentation Jira REST API : https://developer.atlassian.com/cloud/jira/platform/rest/v3/
2. Consultez les logs de l'application pour plus de détails sur les erreurs
3. Vérifiez que votre version de Jira est compatible (9.12.24 ou supérieure recommandée)
