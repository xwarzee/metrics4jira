# Guide de Dépannage - Jira JQL Explorer Excel

Solutions aux problèmes courants rencontrés avec l'application Excel VBA.

## Erreur 404 : "null for uri" lors de la recherche

### Symptôme
```
Jira API request failed: 404
{"message":"null for uri: https://your-server.com/jira/jira/rest/api/2/search","status-code":404}
```

### Causes Possibles

1. **URL en double** : L'URL contient `/jira` deux fois
2. **Slash final dans l'URL** : L'URL se termine par `/`
3. **Mauvais endpoint** : Le chemin de l'API est incorrect

### Solutions

#### Solution 1 : Vérifier l'URL Jira (Recommandé)

Dans la feuille **Config**, vérifiez le format de votre URL :

**✅ CORRECT - Jira Server :**
```
https://your-server.com/jira
```

**❌ INCORRECT :**
```
https://your-server.com/jira/        ← Slash final
https://your-server.com/jira/rest    ← Inclut /rest
https://your-server.com              ← Manque /jira
```

**✅ CORRECT - Jira Cloud :**
```
https://your-domain.atlassian.net
```

**❌ INCORRECT :**
```
https://your-domain.atlassian.net/   ← Slash final
```

#### Solution 2 : Activer le Debug

1. Ouvrez VBA (**Alt + F11**)
2. Appuyez sur **Ctrl + G** pour ouvrir la fenêtre Exécution
3. Lancez une recherche
4. Regardez l'URL affichée dans la fenêtre Exécution :
   ```
   Server API URL: https://your-server.com/jira/rest/api/2/search
   ```
5. Vérifiez que l'URL est correcte

#### Solution 3 : Nettoyer la Configuration

1. Allez dans la feuille **Config**
2. Supprimez le contenu de la cellule **B2** (Jira URL)
3. Re-saisissez l'URL **SANS SLASH À LA FIN**
4. Cliquez sur **Test Connection** pour vérifier

#### Solution 4 : Vérifier l'API Version

Pour **Jira Server on-premise** :
- Sélectionnez : **"Jira Server 9.12.24"** dans API Version
- URL typique : `https://jira.your-company.com/jira`

Pour **Jira Cloud** :
- Sélectionnez : **"Jira Cloud (Current)"** dans API Version
- URL typique : `https://your-company.atlassian.net`

### Vérification Rapide

Testez l'URL dans votre navigateur :

**Pour Jira Server :**
```
https://your-server.com/jira/rest/api/2/myself
```

**Pour Jira Cloud :**
```
https://your-domain.atlassian.net/rest/api/3/myself
```

Vous devriez voir vos informations utilisateur en JSON ou une demande de connexion.

---

## Erreur 403 : XSRF check failed

### Symptôme
```
Jira API request failed: 403
XSRF check failed
```

### Cause
- Jira Server nécessite un header de sécurité anti-CSRF pour les requêtes POST

### Solution

Le code a été corrigé pour inclure automatiquement le header `X-Atlassian-Token: no-check`.

**Vérifiez que votre code contient** (dans `JiraApiClient.bas`, fonction `SearchIssuesServer`) :

```vba
http.setRequestHeader "X-Atlassian-Token", "no-check"
```

Si ce header est manquant :
1. Ouvrez VBA (Alt + F11)
2. Trouvez la fonction `SearchIssuesServer` dans `JiraApiClient`
3. Ajoutez cette ligne avant `http.Send requestBody` :
   ```vba
   http.setRequestHeader "X-Atlassian-Token", "no-check"
   ```
4. Sauvegardez et testez à nouveau

---

## Erreur : Deserialization / VALUE_STRING token

### Symptôme
```
Error deserializing VALUE_STRING token from com.atlassian.jira.rest.v2.search.SearchRequestBean(expand)
```

### Cause
- Le paramètre `expand` dans la requête JSON n'est pas accepté ou mal formaté par Jira Server

### Solution

Le code a été corrigé pour retirer le paramètre `expand` qui n'est pas nécessaire.

**Vérifiez que votre payload JSON** (dans `JiraApiClient.bas`, fonction `SearchIssuesServer`) **ne contient PAS** le paramètre `expand` :

```vba
' CORRECT :
requestBody = "{"
requestBody = requestBody & """jql"":""" & EscapeJson(jql) & ""","
requestBody = requestBody & """startAt"":" & CStr(startAt) & ","
requestBody = requestBody & """maxResults"":" & CStr(maxResults) & ","
requestBody = requestBody & """fields"":[""*all""]"
requestBody = requestBody & "}"

' INCORRECT (avec expand) :
' requestBody = requestBody & """expand"":""names,schema"""  ← À SUPPRIMER
```

Le champ `fields: ["*all"]` suffit pour récupérer tous les champs disponibles.

---

## Erreur 401 : Unauthorized

### Symptôme
```
Jira API request failed: 401
Unauthorized
```

### Causes
- Token API invalide ou expiré
- Username (email) incorrect
- Permissions insuffisantes

### Solutions

1. **Vérifier le Token API**
   - Générez un nouveau token : https://id.atlassian.com/manage-profile/security/api-tokens
   - Copiez le token complet (sans espaces)
   - Collez dans **Config B4**

2. **Vérifier le Username**
   - Utilisez votre **email complet** (ex: `user@company.com`)
   - Pas de nom d'utilisateur ou alias

3. **Vérifier les Permissions**
   - Assurez-vous d'avoir accès au projet Jira
   - Testez d'abord avec une requête simple : `project = YOURPROJECT`

---

## Erreur 400 : Invalid request payload (API v3)

### Symptôme
```
Jira API request failed: 400
{"errorMessages":["invalid request payload"]}
```

### Cause
- Mauvaise API Version sélectionnée pour Jira Cloud

### Solution

1. Allez dans la feuille **Config**
2. Changez **API Version** (cellule B6) à : **"Jira Cloud (Current)"**
3. Testez à nouveau la connexion

**Important** : Jira Cloud utilise API v3 avec GET, pas POST !

---

## Erreur : Cannot parse JSON response

### Symptôme
```
Type mismatch or object required error
```

### Causes
- ScriptControl non disponible (Mac)
- Réponse non-JSON de Jira

### Solutions

#### Sur Mac :
1. Téléchargez VBA-JSON : https://github.com/VBA-tools/VBA-JSON
2. Importez `JsonConverter.bas` dans VBA
3. Modifiez `JiraApiClient.bas` :
   ```vba
   ' Remplacez ParseJson() par :
   Set jsonResponse = JsonConverter.ParseJson(response)
   ```

#### Sur Windows :
1. Vérifiez que **Microsoft Script Control 1.0** est coché dans **Outils > Références**
2. Si manquant, téléchargez-le ou utilisez VBA-JSON (voir Mac)

---

## Erreur : Reference not found / Manquante

### Symptôme
```
Compile error: Can't find project or library
```

### Solution

1. Ouvrez VBA (**Alt + F11**)
2. **Outils** > **Références**
3. Décochez les références marquées **MANQUANT**
4. Cochez ces 3 références :
   - ✅ Microsoft XML, v6.0
   - ✅ Microsoft Scripting Runtime
   - ✅ Microsoft Script Control 1.0 (Windows uniquement)
5. Cliquez **OK**

---

## Erreur : Timeout / Request too long

### Symptôme
```
The operation timed out
```

### Causes
- Trop de résultats
- Requête JQL complexe
- Connexion Internet lente

### Solutions

1. **Réduire Max Results**
   - Changez **Max Results** dans Config à 10 ou 25

2. **Affiner la Requête JQL**
   ```jql
   # Au lieu de :
   project = MYPROJECT

   # Essayez :
   project = MYPROJECT AND created >= -7d
   ```

3. **Paginer les Résultats**
   - Modifiez le code VBA pour utiliser `startAt` et faire plusieurs requêtes

---

## Performance Lente

### Symptômes
- Excel se fige pendant la recherche
- Recherches prennent plusieurs minutes

### Solutions

1. **Réduire le Volume de Données**
   - Max Results ≤ 50
   - Utilisez des filtres JQL précis

2. **Désactiver ScreenUpdating**
   - Le code le fait déjà, mais vérifiez :
   ```vba
   Application.ScreenUpdating = False
   ' ... code ...
   Application.ScreenUpdating = True
   ```

3. **Fermer les autres Feuilles**
   - Gardez seulement Issues et Config ouverts

---

## URL invalide ou connexion refusée

### Symptômes
```
Cannot connect to remote server
Connection refused
```

### Solutions

1. **Vérifier le Protocole**
   - URL doit commencer par `https://`
   - Pas `http://` (non sécurisé)

2. **Vérifier le Réseau**
   - Testez dans un navigateur
   - Vérifiez le VPN si requis
   - Vérifiez le proxy d'entreprise

3. **Pare-feu / Antivirus**
   - Autorisez Excel à faire des connexions HTTPS sortantes
   - Ajoutez l'URL Jira à la liste blanche

---

## Debug Avancé

### Afficher les Logs VBA

1. **Alt + F11** pour ouvrir VBA
2. **Ctrl + G** pour la fenêtre Exécution
3. Les messages `Debug.Print` s'affichent ici :
   ```
   Server API URL: https://...
   Server API payload: {"jql":"..."}
   Cloud API URL: https://...
   ```

### Tester dans la Fenêtre Exécution

```vba
' Charger la config
JiraConfig.LoadConfigFromSheet

' Afficher l'URL
?JiraConfig.Config.JiraUrl

' Afficher l'endpoint complet
?JiraConfig.Config.JiraUrl & JiraConfig.GetSearchEndpoint()

' Tester la connexion
?JiraApiClient.TestConnection()
```

### Ajouter Plus de Logs

Ajoutez dans `JiraApiClient.bas` :

```vba
' Avant http.Send
Debug.Print "Sending request to: " & url
Debug.Print "Authorization: " & JiraConfig.GetAuthHeader()

' Après http.Send
Debug.Print "Response status: " & http.Status
Debug.Print "Response text: " & Left(http.responseText, 500)
```

---

## Checklist de Diagnostic

Avant de chercher de l'aide, vérifiez :

- [ ] URL correcte sans slash final
- [ ] API Version correcte (Server vs Cloud)
- [ ] Token API valide (généré récemment)
- [ ] Username est un email complet
- [ ] Références VBA cochées (XML, Scripting, Script Control)
- [ ] Test dans navigateur fonctionne
- [ ] Max Results raisonnable (≤ 100)
- [ ] Requête JQL simple testée d'abord
- [ ] Logs VBA consultés (Ctrl + G)

---

## Exemples d'URL Correctes

### Jira Cloud

```
✅ https://mycompany.atlassian.net
✅ https://acme.atlassian.net

❌ https://mycompany.atlassian.net/
❌ https://mycompany.atlassian.net/rest
```

### Jira Server

```
✅ https://jira.mycompany.com
✅ https://jira.mycompany.com/jira
✅ https://mycompany.com/jira

❌ https://jira.mycompany.com/jira/
❌ https://jira.mycompany.com/jira/rest
❌ https://mycompany.com (si context path /jira requis)
```

---

## Cas Spéciaux

### Jira Server avec Context Path Personnalisé

Si votre Jira est accessible via `https://server.com/myjira` :

```
Jira URL: https://server.com/myjira
API Version: Jira Server 9.12.24
```

### Jira derrière un Proxy

Configurez le proxy dans Windows :
1. Paramètres Windows > Réseau et Internet > Proxy
2. Configurez le proxy manuel si requis
3. Excel utilisera ces paramètres

### Jira avec Port Personnalisé

```
✅ https://jira.company.com:8443/jira
```

---

## Obtenir de l'Aide

Si le problème persiste :

1. **Collectez les informations** :
   - URL exacte (masquez le domaine si sensible)
   - API Version sélectionnée
   - Message d'erreur complet
   - Logs de la fenêtre Exécution VBA

2. **Testez l'API avec cURL** :
   ```bash
   # Jira Cloud
   curl -u "email@company.com:YOUR_TOKEN" \
     "https://domain.atlassian.net/rest/api/3/search/jql?jql=project=TEST&maxResults=1"

   # Jira Server
   curl -u "email@company.com:YOUR_TOKEN" \
     -H "Content-Type: application/json" \
     -d '{"jql":"project=TEST","maxResults":1}' \
     "https://jira.company.com/jira/rest/api/2/search"
   ```

3. **Consultez les ressources** :
   - [README.md](README.md) - Documentation complète
   - [INSTALLATION.md](INSTALLATION.md) - Guide d'installation
   - [QUICKSTART.md](QUICKSTART.md) - Démarrage rapide
   - Documentation Jira API officielle

---

## Problèmes Connus

### Windows vs Mac
- **ScriptControl** non disponible sur Mac → Utilisez VBA-JSON
- Chemins de fichiers différents

### Excel 2016 vs Versions Plus Récentes
- Compatible Excel 2016 à Excel 365
- Certaines fonctions nécessitent Excel 2016 minimum

### Limitations VBA
- Pas de requêtes asynchrones (Excel se fige pendant l'exécution)
- Limite de 1 million de lignes dans Excel
- Parsing JSON limité sans bibliothèque tierce

---

Bon dépannage ! 🔧
