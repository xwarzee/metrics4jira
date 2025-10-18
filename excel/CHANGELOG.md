# Changelog - Jira JQL Explorer Excel

Historique des modifications et corrections apportées à la version Excel VBA.

## Version 1.2 (2025-10-18)

### ✨ Nouvelles Fonctionnalités

#### Support du Proxy HTTP/HTTPS
**Fonctionnalité** : Ajout de la possibilité de configurer un proxy pour accéder à Jira

**Configuration ajoutée** :
- **Use Proxy** : Activer/désactiver le proxy (Yes/No)
- **Proxy Server** : Adresse du serveur proxy (ex: proxy.company.com)
- **Proxy Port** : Port du proxy (par défaut: 8080)
- **Proxy Username** : Nom d'utilisateur pour l'authentification proxy (optionnel)
- **Proxy Password** : Mot de passe pour l'authentification proxy (optionnel)

**Fichiers modifiés** :
- `JiraConfig.bas` - Ajout des champs proxy dans la configuration
- `JiraApiClient.bas` - Ajout de la fonction `ConfigureProxy()` et utilisation de `ServerXMLHTTP` au lieu de `XMLHTTP`

**Code ajouté** dans `JiraApiClient.bas` :
```vba
Private Sub ConfigureProxy(http As Object)
    Dim proxyUrl As String
    proxyUrl = JiraConfig.Config.ProxyServer & ":" & CStr(JiraConfig.Config.ProxyPort)
    http.setProxy 2, proxyUrl  ' 2 = SXH_PROXY_SET_PROXY

    If Len(JiraConfig.Config.ProxyUsername) > 0 Then
        http.setProxyCredentials JiraConfig.Config.ProxyUsername, JiraConfig.Config.ProxyPassword
    End If
End Sub
```

**Utilisation** :
1. Dans la feuille Config, définir **Use Proxy** à "Yes"
2. Renseigner le serveur proxy et le port
3. (Optionnel) Renseigner les identifiants si le proxy nécessite une authentification
4. Tester la connexion

**Cas d'usage** :
- Environnements d'entreprise avec proxy obligatoire
- Connexions via VPN avec proxy
- Proxies avec ou sans authentification

---

## Version 1.1 (2025-10-18)

### 🐛 Corrections de Bugs

#### Erreur 404 : URL en double
**Problème** : L'URL Jira contenait `/jira` en double dans le chemin, résultant en :
```
https://server.com/jira/jira/rest/api/2/search
```

**Correction** :
- Ajout d'une fonction de nettoyage dans `JiraConfig.LoadConfigFromSheet()`
- Suppression automatique du slash final de l'URL
- Mise à jour des instructions dans la feuille Config

**Fichier modifié** : `JiraConfig.bas`

**Code ajouté** :
```vba
' Remove trailing slash from URL if present
If Right(Config.JiraUrl, 1) = "/" Then
    Config.JiraUrl = Left(Config.JiraUrl, Len(Config.JiraUrl) - 1)
End If
```

---

#### Erreur 403 : XSRF check failed
**Problème** : Jira Server nécessite un header anti-CSRF pour accepter les requêtes POST

**Correction** :
- Ajout du header `X-Atlassian-Token: no-check` dans toutes les requêtes POST
- Appliqué à la fonction `SearchIssuesServer()`

**Fichier modifié** : `JiraApiClient.bas`

**Code ajouté** :
```vba
http.setRequestHeader "X-Atlassian-Token", "no-check"
```

**Référence** : [Atlassian XSRF Protection](https://developer.atlassian.com/server/jira/platform/jira-rest-api-example-oauth-authentication-6291692/)

---

#### Erreur : Deserialization / VALUE_STRING token
**Problème** : Le paramètre `expand: "names,schema"` dans la requête JSON causait une erreur de désérialisation dans Jira Server 9.12.24

**Message d'erreur** :
```
Error deserializing VALUE_STRING token from com.atlassian.jira.rest.v2.search.SearchRequestBean(expand)
ligne 1, colonne 93
```

**Correction** :
- Suppression complète du paramètre `expand` de la requête JSON
- Le paramètre `fields: ["*all"]` suffit pour récupérer tous les champs

**Fichier modifié** : `JiraApiClient.bas`

**Avant** :
```vba
requestBody = requestBody & """fields"":[""*all""],"
requestBody = requestBody & """expand"":""names,schema"""
```

**Après** :
```vba
requestBody = requestBody & """fields"":[""*all""]"
```

---

### 📝 Améliorations de la Documentation

#### Nouveau fichier : TROUBLESHOOTING.md
- Guide complet de dépannage
- Solutions pour erreurs 404, 403, 401, 400
- Erreur de désérialisation JSON
- Exemples d'URLs correctes
- Checklist de diagnostic
- Debug avancé avec VBA

#### Nouveau fichier : CHANGELOG.md (ce fichier)
- Historique des corrections
- Détails techniques des changements
- Références aux fichiers modifiés

#### Mises à jour : README.md
- Tableau des différences API v2/v3 mis à jour
- Ajout du header XSRF dans la documentation
- Suppression de la mention du paramètre `expand`

#### Mises à jour : INSTALLATION.md
- Instructions clarifiées pour le format d'URL
- Exemples avec/sans context path

---

### 🔍 Ajout de Logs de Debug

**Fichier modifié** : `JiraApiClient.bas`

**Logs ajoutés** :
```vba
Debug.Print "Server API URL: " & url
Debug.Print "Server API payload: " & requestBody
Debug.Print "Cloud API URL: " & url & "?" & params
```

**Utilisation** :
1. Ouvrir VBA (Alt + F11)
2. Afficher la fenêtre Exécution (Ctrl + G)
3. Lancer une recherche
4. Observer les URLs et payloads générés

---

## Version 1.0 (2025-10-18)

### 🎉 Version Initiale

#### Fonctionnalités
- Support Jira Server 9.12.24 (API v2)
- Support Jira Cloud (API v3)
- Interface Excel avec 3 feuilles (Config, Issues, FieldExplorer)
- Configuration dans Excel
- Test de connexion
- Recherche JQL
- Explorateur de champs
- Métadonnées des champs
- Compatible Windows et Mac (avec VBA-JSON)

#### Modules VBA
- `JiraConfig.bas` - Gestion de la configuration
- `JiraApiClient.bas` - Client API REST
- `JiraExplorer.bas` - Interface Excel

#### Documentation
- README.md - Documentation complète
- INSTALLATION.md - Guide d'installation
- QUICKSTART.md - Démarrage rapide

---

## Problèmes Résolus par Version

### v1.1
- ✅ Erreur 404 avec URL en double
- ✅ Erreur 403 XSRF check failed
- ✅ Erreur de désérialisation JSON (expand)
- ✅ Documentation de dépannage complète

### v1.0
- ✅ Implémentation initiale
- ✅ Support dual API v2/v3
- ✅ Interface Excel fonctionnelle

---

## Problèmes Connus

### Limitations VBA
- **Parsing JSON** : Utilise ScriptControl (Windows) ou nécessite VBA-JSON (Mac)
- **Synchrone seulement** : Excel se fige pendant les requêtes
- **Pas de pagination automatique** : Limité par Max Results

### Workarounds
- **Mac** : Utiliser VBA-JSON pour le parsing JSON
- **Performance** : Réduire Max Results, affiner les requêtes JQL
- **Pagination** : Modifier le code pour appeler SearchIssues en boucle avec startAt

---

## Compatibilité

### Testé avec
- ✅ Excel 2016 (Windows)
- ✅ Excel 2019 (Windows)
- ✅ Excel 365 (Windows)
- ✅ Jira Server 9.12.24
- ✅ Jira Cloud (API v3)

### Nécessite adaptation
- ⚠️ Excel pour Mac (ScriptControl non disponible → utiliser VBA-JSON)
- ❌ Excel Online (VBA non supporté)

---

## Prochaines Améliorations Possibles

### Court terme
- [ ] Bouton pour afficher les détails d'une issue sélectionnée
- [ ] Gestion de la pagination automatique
- [ ] Cache des métadonnées de champs
- [ ] Historique des requêtes JQL

### Moyen terme
- [ ] Export des résultats vers CSV
- [ ] Graphiques et statistiques
- [ ] Filtres sur les résultats
- [ ] Sauvegarde de requêtes favorites

### Long terme
- [ ] Support d'autres API Jira (sprints, boards, etc.)
- [ ] Création/modification d'issues
- [ ] Gestion des pièces jointes
- [ ] Synchronisation bidirectionnelle

---

## Migration depuis v1.0 vers v1.1

### Étapes
1. **Sauvegarder** votre classeur Excel actuel
2. **Exporter** votre configuration (copier les valeurs de la feuille Config)
3. **Supprimer** les anciens modules VBA :
   - Supprimer `JiraConfig`
   - Supprimer `JiraApiClient`
4. **Importer** les nouveaux modules :
   - Importer `JiraConfig.bas` (v1.1)
   - Importer `JiraApiClient.bas` (v1.1)
   - Garder `JiraExplorer.bas` (inchangé)
5. **Restaurer** votre configuration dans la feuille Config
6. **Tester** la connexion

### Changements non-rétrocompatibles
Aucun - La v1.1 est 100% compatible avec les classeurs créés en v1.0

---

## Support

### Ressources
- [README.md](README.md) - Documentation
- [INSTALLATION.md](INSTALLATION.md) - Installation
- [QUICKSTART.md](QUICKSTART.md) - Démarrage rapide
- [TROUBLESHOOTING.md](TROUBLESHOOTING.md) - Dépannage

### Rapporter un Bug
Pour rapporter un bug, incluez :
1. Version d'Excel
2. Type de Jira (Server 9.12.24 ou Cloud)
3. Message d'erreur complet
4. Logs VBA (fenêtre Exécution)
5. Payload JSON généré (si applicable)

---

## Contributeurs

### Version 1.1
- Correction erreur 404 (URL)
- Correction erreur 403 (XSRF)
- Correction erreur désérialisation (expand)
- Documentation de dépannage

### Version 1.0
- Implémentation initiale
- Support dual API
- Interface Excel

---

## Licence

Ce projet est fourni à des fins éducatives et de démonstration.

---

**Dernière mise à jour** : 18 octobre 2025
