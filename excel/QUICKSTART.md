# Guide de Démarrage Rapide - Jira JQL Explorer Excel

Guide rapide pour commencer à utiliser Jira JQL Explorer dans Excel en 5 minutes.

## Installation en 5 minutes

### 1. Créer le classeur (1 min)
```
1. Ouvrez Excel
2. Fichier > Enregistrer sous
3. Nom : JiraExplorer.xlsm
4. Type : Classeur Excel avec macros (*.xlsm)
5. Enregistrez
```

### 2. Importer les modules VBA (2 min)
```
1. Appuyez sur Alt + F11 (ou Fn + Option + F11 sur Mac)
2. Fichier > Importer un fichier
3. Importez dans l'ordre :
   - JiraConfig.bas
   - JiraApiClient.bas
   - JiraExplorer.bas
```

### 3. Ajouter les références (1 min)
```
1. Dans VBA : Outils > Références
2. Cochez :
   ✅ Microsoft XML, v6.0
   ✅ Microsoft Scripting Runtime
   ✅ Microsoft Script Control 1.0 (Windows)
3. OK
```

### 4. Créer les boutons (1 min)
```
1. Fermez VBA (Alt + Q)
2. Dans Excel : Onglet Développeur > Insérer > Bouton

Créez 4 boutons avec ces macros :
┌─────────────────────────┬────────────────────────────────────┐
│ Texte du bouton         │ Macro assignée                     │
├─────────────────────────┼────────────────────────────────────┤
│ Initialize Workbook     │ JiraExplorer.InitializeWorkbook    │
│ Configure Connection    │ JiraExplorer.ConfigureJiraConnection│
│ Test Connection         │ JiraExplorer.TestJiraConnection    │
│ Search Jira Issues      │ JiraExplorer.SearchJiraIssues      │
└─────────────────────────┴────────────────────────────────────┘
```

## Configuration en 3 étapes

### Étape 1 : Initialiser
```
Cliquez sur : [Initialize Workbook]
→ 3 feuilles créées : Config, Issues, FieldExplorer
```

### Étape 2 : Configurer
```
Allez dans la feuille Config et remplissez :

┌──────────────────┬─────────────────────────────────────────┐
│ Champ            │ Exemple                                 │
├──────────────────┼─────────────────────────────────────────┤
│ Jira URL         │ https://votre-domaine.atlassian.net     │
│ API Version      │ Jira Cloud (Current) ▼                  │
│ Username (Email) │ votreemail@example.com                  │
│ API Token        │ [Votre token depuis Atlassian]          │
│ Max Results      │ 50                                      │
└──────────────────┴─────────────────────────────────────────┘

Obtenir un token API :
→ https://id.atlassian.com/manage-profile/security/api-tokens
→ Create API token
→ Copiez et collez dans Config
```

### Étape 3 : Tester
```
Cliquez sur : [Test Connection]
→ ✅ "Successfully connected to Jira!" = OK
→ ❌ "Failed to connect" = Vérifiez la config
```

## Utilisation Simple

### Rechercher des issues

```
1. Cliquez : [Search Jira Issues]

2. Entrez votre JQL :
   project = MYPROJECT

3. Résultats dans la feuille "Issues"
```

### Exemples JQL Rapides

```jql
# Mon travail en cours
assignee = currentUser() AND status = "In Progress"

# Issues urgentes
priority = Highest AND status != Done

# Créées cette semaine
created >= startOfWeek()

# Bugs ouverts
type = Bug AND status = Open

# Mon backlog
assignee = currentUser() AND status = "To Do" ORDER BY priority DESC
```

### Voir les détails d'une issue

```
Méthode 1 (Manuelle) :
1. Notez le numéro de ligne dans "Issues" (ex: ligne 5)
2. Alt + F11 (ouvrir VBA)
3. Ctrl + G (fenêtre Exécution)
4. Tapez : JiraExplorer.ShowIssueDetails 5
5. Entrée

Méthode 2 (Bouton personnalisé) :
1. Créez un bouton "Show Details"
2. Assignez la macro avec un InputBox pour demander le numéro
```

## Raccourcis Clavier Utiles

### Navigation Excel
```
Alt + F11     → Ouvrir/fermer VBA
Alt + Q       → Quitter VBA
Ctrl + S      → Sauvegarder
```

### Dans VBA
```
F5            → Exécuter une macro
F8            → Exécution pas-à-pas
Ctrl + G      → Fenêtre Exécution
Ctrl + R      → Explorateur de projet
```

## Résolution Rapide des Problèmes

### ❌ "Référence manquante"
```
Solution :
1. Alt + F11
2. Outils > Références
3. Décochez les refs MANQUANTES
4. Recochez la bonne version
```

### ❌ "Cannot connect to Jira"
```
Checklist :
□ URL correcte avec https://
□ Token API valide
□ Bonne API Version sélectionnée
□ Connexion Internet OK
□ Pare-feu autorise HTTPS sortant
```

### ❌ "Invalid request payload"
```
Solution :
→ Assurez-vous d'avoir sélectionné :
  "Jira Cloud (Current)" pour Jira Cloud
  "Jira Server 9.12.24" pour Jira on-premise
```

### ❌ "ScriptControl not found" (Mac)
```
Solution :
1. Téléchargez VBA-JSON :
   https://github.com/VBA-tools/VBA-JSON
2. Importez JsonConverter.bas
3. Modifiez JiraApiClient pour utiliser JsonConverter
```

## API Version : Quelle Choisir ?

```
┌─────────────────────┬──────────────────────────────────────┐
│ Si vous utilisez... │ Choisissez...                        │
├─────────────────────┼──────────────────────────────────────┤
│ Jira Cloud          │ "Jira Cloud (Current)" → API v3      │
│ (*.atlassian.net)   │ Méthode : GET                        │
│                     │ Endpoint : /rest/api/3/search/jql    │
├─────────────────────┼──────────────────────────────────────┤
│ Jira Server         │ "Jira Server 9.12.24" → API v2       │
│ (on-premise)        │ Méthode : POST                       │
│ Version 9.12.24     │ Endpoint : /rest/api/2/search        │
└─────────────────────┴──────────────────────────────────────┘
```

## Astuces Pro

### 1. Sauvegarder plusieurs configs
```
Créez plusieurs feuilles Config :
- Config_Prod
- Config_Test
- Config_Demo

Copiez/collez la feuille Config et modifiez JiraConfig.bas
pour charger la bonne feuille.
```

### 2. Automatiser les recherches
```vba
' Créez une macro personnalisée
Sub SearchMyTasks()
    JiraConfig.LoadConfigFromSheet
    Dim issues As Collection
    Set issues = JiraApiClient.SearchIssues("assignee = currentUser()")
    ' Traiter les issues...
End Sub
```

### 3. Exporter vers d'autres feuilles
```vba
' Copiez les résultats vers une autre feuille
Sub ExportToReport()
    Worksheets("Issues").Range("A1:F100").Copy
    Worksheets("Report").Range("A1").PasteSpecial
End Sub
```

### 4. Créer des rapports automatiques
```vba
' Exécutez plusieurs recherches et compilez
Sub WeeklyReport()
    Call SearchHighPriorityBugs
    Call SearchMyInProgress
    Call SearchBlockedIssues
    Call FormatReport
End Sub
```

## Cheat Sheet JQL

### Opérateurs de base
```jql
=          # Égal
!=         # Différent
<, <=, >, >= # Comparaison
IN         # Dans liste
NOT IN     # Pas dans liste
~          # Contient (texte)
!~         # Ne contient pas
IS EMPTY   # Vide
IS NOT EMPTY # Non vide
```

### Fonctions date
```jql
now()           # Maintenant
startOfDay()    # Début du jour
endOfDay()      # Fin du jour
startOfWeek()   # Début de semaine
startOfMonth()  # Début du mois
-7d             # Il y a 7 jours
-2w             # Il y a 2 semaines
```

### Fonctions utilisateur
```jql
currentUser()   # Utilisateur actuel
membersOf("team") # Membres d'un groupe
```

### Tri
```jql
ORDER BY priority DESC
ORDER BY created ASC
ORDER BY updated DESC, priority ASC
```

## Support

### Ressources
- [Installation complète](INSTALLATION.md)
- [Documentation complète](README.md)
- [API Jira Cloud](https://developer.atlassian.com/cloud/jira/platform/rest/v3/)
- [Guide JQL](https://support.atlassian.com/jira-software-cloud/docs/what-is-advanced-searching-in-jira-cloud/)

### Obtenir de l'aide
1. Consultez la section Dépannage dans README.md
2. Vérifiez la configuration dans la feuille Config
3. Testez avec des requêtes JQL simples d'abord
4. Activez Debug.Print dans VBA pour voir les logs

## Checklist de Validation

Avant de commencer à utiliser :

```
Installation
□ Classeur enregistré en .xlsm
□ 3 modules VBA importés
□ 3 références VBA cochées
□ 4 boutons créés et assignés

Configuration
□ Feuilles Config, Issues, FieldExplorer créées
□ URL Jira renseignée
□ API Version correcte sélectionnée
□ Username (email) renseigné
□ API Token généré et renseigné
□ Test Connection réussi ✅

Prêt à utiliser !
□ Première recherche JQL réussie
□ Résultats affichés dans Issues
□ Détails visibles dans FieldExplorer
```

## Prochaines Étapes

Une fois familiarisé avec les bases :

1. **Personnalisez** les feuilles Excel selon vos besoins
2. **Créez des macros** personnalisées pour vos workflows
3. **Automatisez** les rapports récurrents
4. **Intégrez** avec d'autres données Excel
5. **Partagez** le classeur avec votre équipe (sans le token !)

## Exemple Complet de Workflow

```
Scénario : Rapport hebdomadaire de mes tâches

1. [Initialize Workbook]
2. Configurer dans feuille Config
3. [Test Connection] → ✅
4. [Search Jira Issues]
   → JQL : assignee = currentUser() AND created >= -7d
5. Copier les résultats vers feuille "Rapport Hebdo"
6. Ajouter graphiques et analyses Excel
7. Enregistrer et partager (sans token Config !)

Temps total : 10 minutes
```

Bonne utilisation ! 🚀
