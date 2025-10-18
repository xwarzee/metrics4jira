Attribute VB_Name = "JiraApiClient"
Option Explicit

' ==========================================
' Module: JiraApiClient
' Description: API client for Jira REST API v2 and v3
' Compatible with: Excel 2016+
' Dependencies: Microsoft XML, v6.0
' ==========================================

' ==========================================
' Function: TestConnection
' Description: Test connection to Jira instance
' Returns: Boolean - True if successful, False otherwise
' ==========================================
Public Function TestConnection() As Boolean
    Dim http As Object
    Dim url As String

    On Error GoTo ErrorHandler

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    url = JiraConfig.Config.JiraUrl & JiraConfig.GetApiPath() & "/myself"

    ' Configure proxy if enabled
    If JiraConfig.Config.UseProxy Then
        Call ConfigureProxy(http)
    End If

    http.Open "GET", url, False
    http.setRequestHeader "Authorization", JiraConfig.GetAuthHeader()
    http.setRequestHeader "Accept", "application/json"
    http.Send

    TestConnection = (http.Status = 200)

    Set http = Nothing
    Exit Function

ErrorHandler:
    TestConnection = False
    Set http = Nothing
End Function

' ==========================================
' Function: SearchIssues
' Description: Execute JQL query and return issues
' Parameters:
'   jql - JQL query string
'   startAt - Starting index for pagination (default 0)
'   maxResults - Maximum results to return (default from config)
' Returns: Collection of issue dictionaries
' ==========================================
Public Function SearchIssues(ByVal jql As String, _
                            Optional ByVal startAt As Integer = 0, _
                            Optional ByVal maxResults As Integer = -1) As Collection

    If maxResults = -1 Then maxResults = JiraConfig.Config.MaxResults

    ' Route to appropriate API version
    If JiraConfig.Config.ApiVersionValue = JiraConfig.CLOUD_CURRENT Then
        Set SearchIssues = SearchIssuesCloud(jql, startAt, maxResults)
    Else
        Set SearchIssues = SearchIssuesServer(jql, startAt, maxResults)
    End If
End Function

' ==========================================
' Function: SearchIssuesCloud
' Description: Search using Jira Cloud API v3 (GET request)
' Parameters: Same as SearchIssues
' Returns: Collection of issue dictionaries
' ==========================================
Private Function SearchIssuesCloud(ByVal jql As String, _
                                   ByVal startAt As Integer, _
                                   ByVal maxResults As Integer) As Collection

    Dim http As Object
    Dim url As String
    Dim response As String
    Dim jsonResponse As Object

    On Error GoTo ErrorHandler

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Configure proxy if enabled
    If JiraConfig.Config.UseProxy Then
        Call ConfigureProxy(http)
    End If

    ' Build URL with query parameters
    url = JiraConfig.Config.JiraUrl & JiraConfig.GetSearchEndpoint()
    url = url & "?jql=" & UrlEncode(jql)
    url = url & "&startAt=" & CStr(startAt)
    url = url & "&maxResults=" & CStr(maxResults)
    url = url & "&fields=*navigable"

    Debug.Print "Cloud API URL: " & url

    ' Execute GET request
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", JiraConfig.GetAuthHeader()
    http.setRequestHeader "Accept", "application/json"
    http.Send

    ' Check response
    If http.Status <> 200 Then
        Err.Raise vbObjectError + 1000, "SearchIssuesCloud", _
                  "Jira API request failed: " & http.Status & vbCrLf & http.responseText
    End If

    ' Parse JSON response
    response = http.responseText
    Set jsonResponse = ParseJson(response)

    ' Extract issues
    Set SearchIssuesCloud = ExtractIssues(jsonResponse, response)

    Set http = Nothing
    Set jsonResponse = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SearchIssuesCloud", Err.Description
End Function

' ==========================================
' Function: SearchIssuesServer
' Description: Search using Jira Server API v2 (POST request)
' Parameters: Same as SearchIssues
' Returns: Collection of issue dictionaries
' ==========================================
Private Function SearchIssuesServer(ByVal jql As String, _
                                    ByVal startAt As Integer, _
                                    ByVal maxResults As Integer) As Collection

    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    Dim jsonResponse As Object

    On Error GoTo ErrorHandler

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Configure proxy if enabled
    If JiraConfig.Config.UseProxy Then
        Call ConfigureProxy(http)
    End If

    url = JiraConfig.Config.JiraUrl & JiraConfig.GetSearchEndpoint()

    Debug.Print "Server API URL: " & url

    ' Build JSON request body
    requestBody = "{"
    requestBody = requestBody & """jql"":""" & EscapeJson(jql) & ""","
    requestBody = requestBody & """startAt"":" & CStr(startAt) & ","
    requestBody = requestBody & """maxResults"":" & CStr(maxResults) & ","
    requestBody = requestBody & """fields"":[""*all""]"
    requestBody = requestBody & "}"

    Debug.Print "Server API payload: " & requestBody

    ' Execute POST request
    http.Open "POST", url, False
    http.setRequestHeader "Authorization", JiraConfig.GetAuthHeader()
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "X-Atlassian-Token", "no-check"
    http.Send requestBody

    ' Check response
    Debug.Print "Response Status: " & http.Status

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 1000, "SearchIssuesServer", _
                  "Jira API request failed: " & http.Status & vbCrLf & http.responseText
    End If

    ' Parse JSON response
    response = http.responseText
    Debug.Print "Response length: " & Len(response) & " characters"
    Debug.Print "Response preview: " & Left(response, 500)

    Set jsonResponse = ParseJson(response)

    ' Extract issues - pass the response string too for alternative parsing
    Set SearchIssuesServer = ExtractIssues(jsonResponse, response)

    Debug.Print "Issues found: " & SearchIssuesServer.Count

    Set http = Nothing
    Set jsonResponse = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SearchIssuesServer", Err.Description
End Function

' ==========================================
' Function: GetFieldMetadata
' Description: Get field metadata for display names
' Returns: Dictionary of field IDs to names
' ==========================================
Public Function GetFieldMetadata() As Object
    Dim http As Object
    Dim url As String
    Dim response As String
    Dim jsonResponse As Object
    Dim fieldDict As Object
    Dim field As Object

    On Error GoTo ErrorHandler

    Set fieldDict = CreateObject("Scripting.Dictionary")
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Configure proxy if enabled
    If JiraConfig.Config.UseProxy Then
        Call ConfigureProxy(http)
    End If

    url = JiraConfig.Config.JiraUrl & JiraConfig.GetApiPath() & "/field"

    http.Open "GET", url, False
    http.setRequestHeader "Authorization", JiraConfig.GetAuthHeader()
    http.setRequestHeader "Accept", "application/json"
    http.Send

    If http.Status = 200 Then
        response = http.responseText
        Set jsonResponse = ParseJson(response)

        ' Build dictionary of field ID to name
        If Not jsonResponse Is Nothing Then
            For Each field In jsonResponse
                If Not field Is Nothing Then
                    fieldDict(field("id")) = field("name")
                End If
            Next field
        End If
    End If

    Set GetFieldMetadata = fieldDict

    Set http = Nothing
    Set jsonResponse = Nothing
    Exit Function

ErrorHandler:
    Set GetFieldMetadata = fieldDict
    Set http = Nothing
End Function

' ==========================================
' Function: ExtractIssues
' Description: Extract issues array from JSON response
' Parameters: jsonResponse - Parsed JSON object
' Returns: Collection of issue dictionaries
' ==========================================
Private Function ExtractIssues(jsonResponse As Object, Optional jsonString As String = "") As Collection
    Dim issues As Collection
    Dim issuesArray As Object
    Dim issue As Object
    Dim i As Long
    Dim scriptControl As Object

    Set issues = New Collection

    On Error Resume Next

    If Not jsonResponse Is Nothing Then
        Debug.Print "jsonResponse is not Nothing"

        ' Try to access issues property
        Set issuesArray = jsonResponse.issues

        If Err.Number <> 0 Then
            Debug.Print "Error accessing jsonResponse.issues: " & Err.Description
            Err.Clear
            ' Try with parentheses notation
            Set issuesArray = jsonResponse("issues")
        End If

        If Not issuesArray Is Nothing Then
            Debug.Print "issuesArray is not Nothing"
            Debug.Print "issuesArray type: " & TypeName(issuesArray)

            ' Create a new ScriptControl to access array elements properly
            Set scriptControl = CreateObject("ScriptControl")
            scriptControl.Language = "JScript"

            ' Add the full JSON response to ScriptControl
            scriptControl.AddCode "var jiraResponse = " & jsonString & ";"

            ' Add a simple stringify function for older JScript
            scriptControl.AddCode "function getIssue(index) { return jiraResponse.issues[index]; }"

            ' Get the count using JScript
            On Error Resume Next
            i = scriptControl.Eval("jiraResponse.issues.length")
            Debug.Print "Array length from JScript: " & i

            If Err.Number = 0 And i > 0 Then
                ' Extract each issue directly as an object
                Dim j As Long
                For j = 0 To i - 1
                    Err.Clear
                    Set issue = Nothing

                    ' Get the issue object directly from ScriptControl
                    Set issue = scriptControl.Run("getIssue", j)

                    If Err.Number = 0 And Not issue Is Nothing Then
                        issues.Add issue
                        Debug.Print "Successfully added issue " & j
                    Else
                        Debug.Print "Error getting issue " & j & ": " & Err.Description & " (Err: " & Err.Number & ")"
                    End If
                Next j
            Else
                Debug.Print "Could not get array length or array is empty. Error: " & Err.Description
            End If

            Set scriptControl = Nothing
        Else
            Debug.Print "issuesArray is Nothing"
        End If
    Else
        Debug.Print "jsonResponse is Nothing"
    End If

    On Error GoTo 0

    Set ExtractIssues = issues
End Function

' ==========================================
' Function: ParseJson
' Description: Parse JSON string to object
' Parameters: jsonString - JSON string to parse
' Returns: Parsed JSON object
' ==========================================
Private Function ParseJson(ByVal jsonString As String) As Object
    Dim scriptControl As Object

    On Error GoTo ErrorHandler

    Set scriptControl = CreateObject("ScriptControl")
    scriptControl.Language = "JScript"

    ' Use JScript to parse JSON
    scriptControl.AddCode "function parseJson(json) { return eval('(' + json + ')'); }"
    Set ParseJson = scriptControl.Run("parseJson", jsonString)

    Set scriptControl = Nothing
    Exit Function

ErrorHandler:
    Set ParseJson = Nothing
    Set scriptControl = Nothing
End Function

' ==========================================
' Function: HasKey
' Description: Check if object has key
' Parameters:
'   obj - Object to check
'   key - Key to look for
' Returns: Boolean - True if key exists
' ==========================================
Private Function HasKey(obj As Object, key As String) As Boolean
    On Error Resume Next
    HasKey = Not IsEmpty(obj(key))
    On Error GoTo 0
End Function

' ==========================================
' Function: UrlEncode
' Description: URL encode a string
' Parameters: text - String to encode
' Returns: URL encoded string
' ==========================================
Private Function UrlEncode(ByVal text As String) As String
    Dim i As Integer
    Dim char As String
    Dim result As String

    result = ""
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        Select Case char
            Case " "
                result = result & "%20"
            Case "!"
                result = result & "%21"
            Case "#"
                result = result & "%23"
            Case "$"
                result = result & "%24"
            Case "&"
                result = result & "%26"
            Case "'"
                result = result & "%27"
            Case "("
                result = result & "%28"
            Case ")"
                result = result & "%29"
            Case "*"
                result = result & "%2A"
            Case "+"
                result = result & "%2B"
            Case ","
                result = result & "%2C"
            Case "/"
                result = result & "%2F"
            Case ":"
                result = result & "%3A"
            Case ";"
                result = result & "%3B"
            Case "="
                result = result & "%3D"
            Case "?"
                result = result & "%3F"
            Case "@"
                result = result & "%40"
            Case "["
                result = result & "%5B"
            Case "]"
                result = result & "%5D"
            Case Else
                result = result & char
        End Select
    Next i

    UrlEncode = result
End Function

' ==========================================
' Function: EscapeJson
' Description: Escape string for JSON
' Parameters: text - String to escape
' Returns: Escaped string
' ==========================================
Private Function EscapeJson(ByVal text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeJson = result
End Function

' ==========================================
' Function: ConfigureProxy
' Description: Configure proxy settings for HTTP object
' Parameters: http - XMLHTTP object
' ==========================================
Private Sub ConfigureProxy(http As Object)
    Dim proxyUrl As String

    ' Build proxy URL
    proxyUrl = JiraConfig.Config.ProxyServer & ":" & CStr(JiraConfig.Config.ProxyPort)

    Debug.Print "Configuring proxy: " & proxyUrl

    ' Set proxy server
    http.setProxy 2, proxyUrl  ' 2 = SXH_PROXY_SET_PROXY

    ' Set proxy credentials if provided
    If Len(JiraConfig.Config.ProxyUsername) > 0 Then
        http.setProxyCredentials JiraConfig.Config.ProxyUsername, JiraConfig.Config.ProxyPassword
        Debug.Print "Proxy credentials configured for user: " & JiraConfig.Config.ProxyUsername
    End If
End Sub
