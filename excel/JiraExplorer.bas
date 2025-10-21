Attribute VB_Name = "JiraExplorer"
Option Explicit

' ==========================================
' Module: JiraExplorer
' Description: Main interface module for Jira JQL Explorer
' Compatible with: Excel 2016+
' ==========================================

Private fieldMetadata As Object

' ==========================================
' Subroutine: InitializeWorkbook
' Description: Initialize workbook with required sheets
' ==========================================
Public Sub InitializeWorkbook()
    Application.ScreenUpdating = False

    ' Create or get Config sheet
    Call EnsureSheetExists("Config")
    Call JiraConfig.InitializeConfig
    Call JiraConfig.LoadConfigFromSheet

    ' Create or get Issues sheet
    Call EnsureSheetExists("Issues")
    Call CreateIssuesSheetLayout

    ' Create or get FieldExplorer sheet
    Call EnsureSheetExists("FieldExplorer")
    Call CreateFieldExplorerLayout

    Application.ScreenUpdating = True

    MsgBox "Workbook initialized!" & vbCrLf & vbCrLf & _
           "1. Configure your Jira connection in the 'Config' sheet" & vbCrLf & _
           "2. Use the 'Search Jira Issues' button to query issues" & vbCrLf & _
           "3. Click on an issue to see details in 'FieldExplorer' sheet", _
           vbInformation, "Jira JQL Explorer"
End Sub

' ==========================================
' Subroutine: ConfigureJiraConnection
' Description: Show configuration dialog
' ==========================================
Public Sub ConfigureJiraConnection()
    Dim ws As Worksheet

    ' Load current configuration
    Call JiraConfig.LoadConfigFromSheet

    ' Activate Config sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Config")
    On Error GoTo 0

    If Not ws Is Nothing Then
        ws.Activate
        MsgBox "Please configure your Jira connection in this sheet." & vbCrLf & vbCrLf & _
               "After configuration, click 'Test Connection' to verify.", _
               vbInformation, "Jira Configuration"
    Else
        MsgBox "Config sheet not found. Please run 'Initialize Workbook' first.", _
               vbCritical, "Error"
    End If
End Sub

' ==========================================
' Subroutine: TestJiraConnection
' Description: Test connection to Jira
' ==========================================
Public Sub TestJiraConnection()
    Dim success As Boolean

    ' Load configuration
    Call JiraConfig.LoadConfigFromSheet

    ' Validate configuration
    If Not JiraConfig.IsConfigValid Then
        MsgBox "Configuration is incomplete. Please fill in all required fields.", _
               vbCritical, "Configuration Error"
        Exit Sub
    End If

    ' Test connection
    Application.Cursor = xlWait
    success = JiraApiClient.TestConnection()
    Application.Cursor = xlDefault

    If success Then
        MsgBox "Successfully connected to Jira!" & vbCrLf & vbCrLf & _
               "URL: " & JiraConfig.Config.JiraUrl & vbCrLf & _
               "API Version: " & IIf(JiraConfig.Config.ApiVersionValue = JiraConfig.CLOUD_CURRENT, _
                                     "Jira Cloud (v3)", "Jira Server 9.12.24 (v2)"), _
               vbInformation, "Connection Successful"

        ' Load field metadata
        Set fieldMetadata = JiraApiClient.GetFieldMetadata()
    Else
        MsgBox "Failed to connect to Jira." & vbCrLf & vbCrLf & _
               "Please check:" & vbCrLf & _
               "- Jira URL is correct" & vbCrLf & _
               "- Username (email) is correct" & vbCrLf & _
               "- API Token is valid" & vbCrLf & _
               "- Network connection is working", _
               vbCritical, "Connection Failed"
    End If
End Sub

' ==========================================
' Subroutine: SearchJiraIssues
' Description: Execute JQL search and display results
' ==========================================
Public Sub SearchJiraIssues()
    Dim jql As String
    Dim issues As Collection
    Dim ws As Worksheet

    ' Load configuration
    Call JiraConfig.LoadConfigFromSheet

    ' Validate configuration
    If Not JiraConfig.IsConfigValid Then
        MsgBox "Configuration is incomplete. Please configure Jira connection first.", _
               vbCritical, "Configuration Error"
        Call ConfigureJiraConnection
        Exit Sub
    End If

    ' Get JQL query from user
    jql = InputBox("Enter JQL Query:" & vbCrLf & vbCrLf & _
                   "Examples:" & vbCrLf & _
                   "  project = MYPROJECT" & vbCrLf & _
                   "  assignee = currentUser()" & vbCrLf & _
                   "  status = Open AND priority = High" & vbCrLf & _
                   "  created >= -7d", _
                   "JQL Query", _
                   "project = MYPROJECT")

    If Len(jql) = 0 Then Exit Sub

    ' Execute search
    Application.ScreenUpdating = False
    Application.Cursor = xlWait

    On Error GoTo ErrorHandler

    Set issues = JiraApiClient.SearchIssues(jql)

    ' Display results
    Set ws = ThisWorkbook.Worksheets("Issues")
    Call DisplayIssues(ws, issues)

    Application.ScreenUpdating = True
    Application.Cursor = xlDefault

    MsgBox "Search completed!" & vbCrLf & vbCrLf & _
           "Found " & issues.Count & " issue(s)." & vbCrLf & vbCrLf & _
           "Click on an issue row to view details.", _
           vbInformation, "Search Results"

    ws.Activate
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    MsgBox "Search failed:" & vbCrLf & vbCrLf & Err.Description, _
           vbCritical, "Search Error"
End Sub

' ==========================================
' Subroutine: DisplayIssues
' Description: Display issues in the Issues sheet
' Parameters:
'   ws - Worksheet to display in
'   issues - Collection of issues
' ==========================================
Private Sub DisplayIssues(ws As Worksheet, issues As Collection)
    Dim issue As Object
    Dim fields As Object
    Dim row As Long

    ' Clear existing data
    ws.Rows("2:" & ws.Rows.Count).ClearContents

    ' Display issues
    row = 2
    For Each issue In issues
        If Not issue Is Nothing Then
            ' Get fields object using Dictionary.Item() for VBA-JSON compatibility
            On Error Resume Next
            Set fields = issue.Item("fields")
            On Error GoTo 0

            If Not fields Is Nothing Then
                Dim epicKey As String

                ws.Cells(row, 1).Value = GetDictValue(issue, "key")
                ws.Cells(row, 2).Value = GetDictValue(fields, "summary")
                ws.Cells(row, 3).Value = GetDictNestedValue(fields, "status", "name")
                ws.Cells(row, 4).Value = GetDictNestedValue(fields, "priority", "name")
                ws.Cells(row, 5).Value = GetDictNestedValue(fields, "assignee", "displayName")

                ' Get Epic Link and Epic Summary
                epicKey = GetEpicLink(fields)
                ws.Cells(row, 6).Value = epicKey
                ws.Cells(row, 7).Value = GetEpicSummary(epicKey)

                ws.Cells(row, 8).Value = GetLabels(fields)
                ws.Cells(row, 9).Value = GetFixVersions(fields)
                Dim sprintValue As String
                sprintValue = GetSprints(fields)
                Debug.Print "Row " & row & " Sprint value: '" & sprintValue & "'"
                ws.Cells(row, 10).Value = sprintValue
                ws.Cells(row, 11).Value = GetDictValue(fields, "created")
            End If

            ' Store full issue data in hidden column for detail view
            ws.Cells(row, 12).Value = row
            ws.Cells(row, 12).NumberFormat = "0"

            row = row + 1
        End If
    Next issue

    ' Auto-fit columns
    ws.Columns("A:K").AutoFit

    ' Update result count
    ws.Range("L1").Value = "Total: " & issues.Count
End Sub

' ==========================================
' Subroutine: ShowIssueDetails
' Description: Show details of selected issue
' Called from worksheet change event
' ==========================================
Public Sub ShowIssueDetails(issueRow As Long)
    Dim wsIssues As Worksheet
    Dim wsExplorer As Worksheet
    Dim issueKey As String
    Dim jql As String
    Dim issues As Collection
    Dim issue As Object
    Dim fields As Object

    Set wsIssues = ThisWorkbook.Worksheets("Issues")
    Set wsExplorer = ThisWorkbook.Worksheets("FieldExplorer")

    ' Get issue key
    issueKey = wsIssues.Cells(issueRow, 1).Value
    If Len(issueKey) = 0 Then Exit Sub

    ' Load configuration
    Call JiraConfig.LoadConfigFromSheet

    Application.ScreenUpdating = False
    Application.Cursor = xlWait

    On Error GoTo ErrorHandler

    ' Load field metadata if not already loaded
    If fieldMetadata Is Nothing Then
        Set fieldMetadata = JiraApiClient.GetFieldMetadata()
    End If

    ' Fetch issue details as JSON string
    Dim issueJson As String
    issueJson = JiraApiClient.GetIssueJson(issueKey)

    If Len(issueJson) > 0 Then
        Debug.Print "Issue JSON length: " & Len(issueJson)
        Debug.Print "Issue JSON preview: " & Left(issueJson, 200)
        ' Display field explorer - pass the JSON string
        Call DisplayFieldExplorerSimple(wsExplorer, issueKey, issueJson)
    Else
        MsgBox "Unable to fetch issue details", vbExclamation
    End If

    Application.ScreenUpdating = True
    Application.Cursor = xlDefault

    wsExplorer.Activate
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    MsgBox "Failed to load issue details:" & vbCrLf & vbCrLf & Err.Description, _
           vbExclamation, "Error"
End Sub

' ==========================================
' Subroutine: DisplayFieldExplorer
' Description: Display all fields of an issue
' Parameters:
'   ws - Worksheet to display in
'   issueKey - Issue key
'   fields - Fields object
' ==========================================
Private Sub DisplayFieldExplorer(ws As Worksheet, issueKey As String, fields As Object)
    Dim row As Long
    Dim key As Variant
    Dim fieldName As String
    Dim fieldValue As String
    Dim scriptControl As Object
    Dim fieldObj As Variant

    ' Clear existing data
    ws.Rows("3:" & ws.Rows.Count).ClearContents

    ' Set issue key
    ws.Range("B1").Value = issueKey

    ' Create ScriptControl for accessing field values
    Set scriptControl = CreateObject("ScriptControl")
    scriptControl.Language = "JScript"
    scriptControl.AddObject "fieldsObj", fields, True

    ' Add helper function to convert object to string (without JSON.stringify)
    scriptControl.AddCode "function objToString(obj) {" & _
        "  if (obj === null || obj === undefined) return 'null';" & _
        "  var str = '{';" & _
        "  var first = true;" & _
        "  for (var k in obj) {" & _
        "    if (!first) str += ', ';" & _
        "    str += k + ': ' + obj[k];" & _
        "    first = false;" & _
        "  }" & _
        "  return str + '}';" & _
        "}"

    ' Add helper function to get field value as string
    scriptControl.AddCode "function getFieldValue(key) {" & _
        "  try {" & _
        "    var val = fieldsObj[key];" & _
        "    if (val === null) return '[null]';" & _
        "    if (val === undefined) return '[undefined]';" & _
        "    var valType = typeof val;" & _
        "    if (valType === 'string') return val;" & _
        "    if (valType === 'number') return String(val);" & _
        "    if (valType === 'boolean') return val ? 'true' : 'false';" & _
        "    if (valType === 'object') {" & _
        "      if (val.name !== undefined) return String(val.name);" & _
        "      if (val.value !== undefined) return String(val.value);" & _
        "      if (val.displayName !== undefined) return String(val.displayName);" & _
        "      if (val.key !== undefined) return String(val.key);" & _
        "      return objToString(val);" & _
        "    }" & _
        "    return '[type: ' + valType + ']';" & _
        "  } catch(e) { return '[Error: ' + e.message + ']'; }" & _
        "}"

    ' Display fields
    row = 3
    Dim keys As Variant
    keys = GetObjectKeys(fields)

    Debug.Print "Number of keys found: " & UBound(keys) - LBound(keys) + 1

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        key = keys(i)

        ' Get field name from metadata
        If Not fieldMetadata Is Nothing Then
            If fieldMetadata.Exists(CStr(key)) Then
                fieldName = fieldMetadata(CStr(key))
            Else
                fieldName = CStr(key)
            End If
        Else
            fieldName = CStr(key)
        End If

        ' Get field value using helper function
        On Error Resume Next
        fieldValue = scriptControl.Run("getFieldValue", CStr(key))
        If Err.Number <> 0 Then
            fieldValue = "[Error: " & Err.Description & "]"
            Debug.Print "Error for key '" & key & "': " & Err.Description
            Err.Clear
        Else
            Debug.Print "Key: " & key & " = " & Left(fieldValue, 50)
        End If
        On Error GoTo 0

        ws.Cells(row, 1).Value = fieldName
        ws.Cells(row, 2).Value = fieldValue

        row = row + 1
    Next i

    ' Auto-fit columns
    ws.Columns("A:B").AutoFit

    Set scriptControl = Nothing
End Sub

' ==========================================
' Subroutine: DisplayFieldExplorerSimple
' Description: Display issue fields in FieldExplorer (simplified version)
' Parameters:
'   ws - Worksheet to display in
'   issueKey - Issue key
'   issueJson - JSON string of the issue or search result
' ==========================================
Private Sub DisplayFieldExplorerSimple(ws As Worksheet, issueKey As String, issueJson As String)
    Dim row As Long
    Dim commonFields As Variant
    Dim i As Long
    Dim fieldKey As String
    Dim fieldName As String
    Dim fieldValue As String
    Dim jsonResponse As Object
    Dim issueObj As Object
    Dim fieldsObj As Object

    On Error Resume Next

    ' Clear existing data
    ws.Rows("3:" & ws.Rows.Count).ClearContents

    ' Set issue key
    ws.Range("B1").Value = issueKey

    ' Parse JSON using VBA-JSON
    Set jsonResponse = JsonConverter.ParseJson(issueJson)

    If jsonResponse Is Nothing Then
        Debug.Print "Failed to parse JSON"
        Exit Sub
    End If

    ' Extract issue object - for Jira Server, the response has issues array; for Cloud, it's the issue directly
    If TypeName(jsonResponse) = "Dictionary" Then
        ' Check if it has an "issues" property (Jira Server response)
        Err.Clear
        Set issueObj = jsonResponse.Item("issues")
        If Err.Number = 0 And Not issueObj Is Nothing Then
            ' It's a search result with issues array
            If TypeName(issueObj) = "Collection" And issueObj.Count > 0 Then
                Set issueObj = issueObj.Item(1)  ' Get first issue
            End If
        Else
            ' It's a direct issue object (Jira Cloud response)
            Set issueObj = jsonResponse
        End If
        Err.Clear
    End If

    If issueObj Is Nothing Then
        Debug.Print "Could not extract issue object"
        Exit Sub
    End If

    ' Get fields object
    Set fieldsObj = issueObj.Item("fields")
    If fieldsObj Is Nothing Then
        Debug.Print "Could not extract fields object"
        Exit Sub
    End If

    Debug.Print "Fields object type: " & TypeName(fieldsObj)

    ' Define common Jira fields to display
    commonFields = Array("summary", "description", "status", "priority", "assignee", _
                         "reporter", "created", "updated", "resolutiondate", "duedate", _
                         "issuetype", "project", "components", "labels", "fixVersions", _
                         "versions", "timeoriginalestimate", "timeestimate", "timespent", _
                         "aggregatetimeoriginalestimate", "aggregatetimeestimate", "aggregatetimespent", _
                         "resolution", "customfield_10102", "customfield_10109", "customfield_10014", "customfield_10008", "customfield_10011")

    ' Display fields
    row = 3
    For i = LBound(commonFields) To UBound(commonFields)
        fieldKey = commonFields(i)

        ' Get field name from metadata or use key
        If Not fieldMetadata Is Nothing And fieldMetadata.Exists(fieldKey) Then
            fieldName = fieldMetadata(fieldKey)
        Else
            fieldName = fieldKey
        End If

        ' Get field value using helper function
        fieldValue = FormatDictFieldValue(fieldsObj, fieldKey)

        ' Only display if field has a value
        If Len(fieldValue) > 0 Then
            Debug.Print "Field '" & fieldKey & "' = '" & Left(fieldValue, 50) & "'"
            ws.Cells(row, 1).Value = fieldName
            ws.Cells(row, 2).Value = fieldValue
            row = row + 1
        End If
    Next i

    ' Add Epic Summary field
    Dim epicKey As String
    epicKey = GetEpicLink(fieldsObj)
    If Len(epicKey) > 0 Then
        fieldValue = GetEpicSummary(epicKey)
        If Len(fieldValue) > 0 Then
            ws.Cells(row, 1).Value = "Epic Summary"
            ws.Cells(row, 2).Value = fieldValue
            row = row + 1
        End If
    End If

    ' Auto-fit columns
    ws.Columns("A:B").AutoFit

    On Error GoTo 0
End Sub

' ==========================================
' Function: GetObjectKeys
' Description: Get keys from object (JScript object)
' Parameters: obj - Object to get keys from
' Returns: Array of keys
' ==========================================
Private Function GetObjectKeys(obj As Object) As Variant
    Dim scriptControl As Object
    Dim keys As Variant
    Dim i As Long
    Dim keyArray() As Variant
    Dim keyCount As Long

    On Error GoTo ErrorHandler

    Debug.Print "GetObjectKeys: Starting..."
    Debug.Print "GetObjectKeys: Object type = " & TypeName(obj)

    Set scriptControl = CreateObject("ScriptControl")
    scriptControl.Language = "JScript"

    ' Add the object to the script context
    On Error Resume Next
    scriptControl.AddObject "fieldsObj", obj, True
    If Err.Number <> 0 Then
        Debug.Print "GetObjectKeys: Error adding object - " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    ' Create function to get keys
    scriptControl.AddCode "function getKeys() { " & _
        "  var arr = new Array(); " & _
        "  var count = 0; " & _
        "  for (var k in fieldsObj) { " & _
        "    arr.push(k); " & _
        "    count++; " & _
        "  } " & _
        "  return arr; " & _
        "}"

    ' Get the keys array
    Set keys = scriptControl.Run("getKeys")
    Debug.Print "GetObjectKeys: Got keys object, type = " & TypeName(keys)

    ' Convert JScript array to VBA array
    On Error Resume Next
    keyCount = keys.length
    Debug.Print "GetObjectKeys: Key count = " & keyCount & ", Err = " & Err.Number

    If Err.Number <> 0 Then
        Debug.Print "GetObjectKeys: Error getting length - " & Err.Description
        GetObjectKeys = Array()
        Exit Function
    End If

    If keyCount = 0 Then
        Debug.Print "GetObjectKeys: No keys found"
        GetObjectKeys = Array()
        Exit Function
    End If

    ReDim keyArray(0 To keyCount - 1)
    For i = 0 To keyCount - 1
        keyArray(i) = keys(i)
        Debug.Print "GetObjectKeys: Key " & i & " = " & keyArray(i)
    Next i

    GetObjectKeys = keyArray
    Debug.Print "GetObjectKeys: Returning " & keyCount & " keys"

    Set scriptControl = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "GetObjectKeys: Error - " & Err.Description
    ' Return empty array on error
    GetObjectKeys = Array()
    Set scriptControl = Nothing
End Function

' ==========================================
' Function: FormatFieldValue
' Description: Format field value for display
' Parameters: value - Value to format
' Returns: String representation
' ==========================================
Private Function FormatFieldValue(value As Variant) As String
    If IsNull(value) Then
        FormatFieldValue = ""
    ElseIf IsObject(value) Then
        FormatFieldValue = "[Object]"
    ElseIf IsArray(value) Then
        FormatFieldValue = "[Array]"
    Else
        FormatFieldValue = CStr(value)
    End If
End Function

' ==========================================
' Function: FormatDictFieldValue
' Description: Format Dictionary field value for display (VBA-JSON compatible)
' Parameters:
'   fieldsObj - Dictionary object containing fields
'   fieldKey - Key of the field to format
' Returns: Formatted string value
' ==========================================
Private Function FormatDictFieldValue(fieldsObj As Object, fieldKey As String) As String
    Dim fieldValue As Variant
    Dim result As String
    Dim item As Variant
    Dim items As Collection

    On Error Resume Next

    ' Get field value
    fieldValue = fieldsObj.Item(fieldKey)

    If Err.Number <> 0 Or IsEmpty(fieldValue) Or IsNull(fieldValue) Then
        FormatDictFieldValue = ""
        Err.Clear
        Exit Function
    End If

    Err.Clear

    ' Format based on type
    If IsObject(fieldValue) Then
        ' Check if it's a Dictionary (single object)
        If TypeName(fieldValue) = "Dictionary" Then
            ' Try common properties
            result = ""

            ' Try "name" property
            On Error Resume Next
            result = fieldValue.Item("name")
            If Err.Number <> 0 Or IsEmpty(result) Then
                Err.Clear
                ' Try "displayName" property
                result = fieldValue.Item("displayName")
                If Err.Number <> 0 Or IsEmpty(result) Then
                    Err.Clear
                    ' Try "value" property
                    result = fieldValue.Item("value")
                    If Err.Number <> 0 Or IsEmpty(result) Then
                        Err.Clear
                        ' Try "key" property
                        result = fieldValue.Item("key")
                        If Err.Number <> 0 Or IsEmpty(result) Then
                            result = ""
                        End If
                    End If
                End If
            End If
            On Error GoTo 0

            FormatDictFieldValue = result

        ' Check if it's a Collection (array)
        ElseIf TypeName(fieldValue) = "Collection" Then
            result = ""
            Set items = fieldValue

            For Each item In items
                If Not IsEmpty(item) And Not IsNull(item) Then
                    If Len(result) > 0 Then result = result & ", "

                    ' If item is a string, add it directly
                    If VarType(item) = vbString Then
                        result = result & item
                    ' If item is a Dictionary, try to get its name/value/key
                    ElseIf TypeName(item) = "Dictionary" Then
                        On Error Resume Next
                        Dim itemValue As String
                        itemValue = item.Item("name")
                        If Err.Number <> 0 Or Len(itemValue) = 0 Then
                            Err.Clear
                            itemValue = item.Item("value")
                            If Err.Number <> 0 Or Len(itemValue) = 0 Then
                                Err.Clear
                                itemValue = item.Item("key")
                                If Err.Number <> 0 Then itemValue = ""
                            End If
                        End If
                        On Error GoTo 0

                        If Len(itemValue) > 0 Then result = result & itemValue
                    Else
                        result = result & CStr(item)
                    End If
                End If
            Next item

            FormatDictFieldValue = result
        Else
            FormatDictFieldValue = ""
        End If
    Else
        ' Simple value (string, number, boolean)
        FormatDictFieldValue = CStr(fieldValue)
    End If

    On Error GoTo 0
End Function

' ==========================================
' Function: GetValue
' Description: Safely get value from object (JScript compatible)
' ==========================================
Private Function GetValue(obj As Object, key As String) As String
    On Error Resume Next

    ' Try direct property access first (for JScript objects)
    Dim result As Variant
    result = CallByName(obj, key, VbGet)

    If Err.Number = 0 Then
        GetValue = CStr(result)
    Else
        GetValue = ""
    End If

    On Error GoTo 0
End Function

' ==========================================
' Function: GetDictValue
' Description: Safely get value from Dictionary object (VBA-JSON compatible)
' ==========================================
Private Function GetDictValue(obj As Object, key As String) As String
    On Error Resume Next

    Dim result As Variant

    ' Use .Item() method for Dictionary objects
    result = obj.Item(key)

    If Err.Number = 0 And Not IsEmpty(result) And Not IsNull(result) Then
        If IsObject(result) Then
            GetDictValue = ""
        Else
            GetDictValue = CStr(result)
        End If
    Else
        GetDictValue = ""
    End If

    On Error GoTo 0
End Function

' ==========================================
' Function: GetNestedValue
' Description: Safely get nested value from object (JScript compatible)
' ==========================================
Private Function GetNestedValue(obj As Object, key1 As String, key2 As String) As String
    Dim nested As Object
    On Error Resume Next

    ' Try to get nested object using CallByName
    Set nested = CallByName(obj, key1, VbGet)

    If Err.Number = 0 And Not nested Is Nothing Then
        Dim result As Variant
        result = CallByName(nested, key2, VbGet)
        If Err.Number = 0 Then
            GetNestedValue = CStr(result)
        Else
            GetNestedValue = ""
        End If
    Else
        GetNestedValue = ""
    End If

    On Error GoTo 0
End Function

' ==========================================
' Function: GetDictNestedValue
' Description: Safely get nested value from Dictionary object (VBA-JSON compatible)
' ==========================================
Private Function GetDictNestedValue(obj As Object, key1 As String, key2 As String) As String
    Dim nested As Object
    On Error Resume Next

    ' Try to get nested Dictionary using .Item() method
    Set nested = obj.Item(key1)

    If Err.Number = 0 And Not nested Is Nothing Then
        Dim result As Variant
        result = nested.Item(key2)
        If Err.Number = 0 And Not IsEmpty(result) And Not IsNull(result) Then
            If IsObject(result) Then
                GetDictNestedValue = ""
            Else
                GetDictNestedValue = CStr(result)
            End If
        Else
            GetDictNestedValue = ""
        End If
    Else
        GetDictNestedValue = ""
    End If

    On Error GoTo 0
End Function

' ==========================================
' Function: GetEpicLink
' Description: Get Epic Link from fields (tries multiple custom field IDs)
' ==========================================
Private Function GetSprints(fields As Object) As String
    Dim sprintsCollection As Object
    Dim sprint As Variant
    Dim sprintNames As String
    Dim sprintName As String

    On Error Resume Next

    ' Try different custom field IDs
    Dim fieldIds As Variant
    fieldIds = Array("customfield_10108", "customfield_10020", "customfield_10010", "customfield_10104", "customfield_10001")

    Dim fieldId As Variant
    For Each fieldId In fieldIds
        Err.Clear
        Set sprintsCollection = Nothing
        Set sprintsCollection = fields.Item(fieldId)

        If Not sprintsCollection Is Nothing Then
            Debug.Print "Found sprint collection in field: " & fieldId
            Debug.Print "Sprint collection type: " & TypeName(sprintsCollection)

            ' VBA-JSON returns Collection for arrays
            If TypeName(sprintsCollection) = "Collection" Then
                sprintNames = ""

                For Each sprint In sprintsCollection
                    sprintName = ""

                    ' Sprints can be either Dictionary objects or strings
                    If TypeName(sprint) = "Dictionary" Then
                        ' Sprint is an object with properties
                        Err.Clear
                        sprintName = sprint.Item("name")
                        If Err.Number <> 0 Then
                            sprintName = ""
                            Err.Clear
                        End If
                    ElseIf VarType(sprint) = vbString Then
                        ' Sprint is a string representation, extract name
                        sprintName = ExtractSprintName(CStr(sprint))
                    End If

                    If Len(sprintName) > 0 Then
                        If Len(sprintNames) > 0 Then
                            sprintNames = sprintNames & ", "
                        End If
                        sprintNames = sprintNames & sprintName
                    End If
                Next sprint

                If Len(sprintNames) > 0 Then
                    GetSprints = sprintNames
                    Debug.Print "Final sprints: " & sprintNames
                    On Error GoTo 0
                    Exit Function
                End If
            End If
        End If
    Next fieldId

    GetSprints = ""
    On Error GoTo 0
End Function

' Helper function to extract sprint name from string
Private Function ExtractSprintName(sprintStr As String) As String
    Dim namePos As Long
    Dim nameStart As Long
    Dim nameEnd As Long
    Dim nameValue As String

    ' Find "name=" in the string
    namePos = InStr(1, sprintStr, "name=", vbTextCompare)

    If namePos > 0 Then
        nameStart = namePos + 5 ' Length of "name="

        ' Check if name is quoted
        If Mid(sprintStr, nameStart, 1) = """" Or Mid(sprintStr, nameStart, 1) = "'" Then
            ' Name is quoted, find closing quote
            Dim quoteChar As String
            quoteChar = Mid(sprintStr, nameStart, 1)
            nameStart = nameStart + 1
            nameEnd = InStr(nameStart, sprintStr, quoteChar)
        Else
            ' Name is not quoted, find next comma or bracket
            nameEnd = InStr(nameStart, sprintStr, ",")
            Dim bracketPos As Long
            bracketPos = InStr(nameStart, sprintStr, "]")

            If bracketPos > 0 And (nameEnd = 0 Or bracketPos < nameEnd) Then
                nameEnd = bracketPos
            End If
        End If

        If nameEnd > nameStart Then
            nameValue = Mid(sprintStr, nameStart, nameEnd - nameStart)
            ExtractSprintName = Trim(nameValue)
        End If
    End If
End Function

Private Function GetFixVersions(fields As Object) As String
    Dim versionsCollection As Object
    Dim version As Variant
    Dim versionName As String
    Dim result As String

    On Error Resume Next

    ' Get fixVersions using .Item()
    Set versionsCollection = fields.Item("fixVersions")

    If Err.Number = 0 And Not versionsCollection Is Nothing Then
        ' VBA-JSON returns Collection for arrays
        If TypeName(versionsCollection) = "Collection" Then
            result = ""
            For Each version In versionsCollection
                If Not version Is Nothing Then
                    ' Each version is a Dictionary
                    If TypeName(version) = "Dictionary" Then
                        versionName = version.Item("name")
                        If Err.Number = 0 And Len(versionName) > 0 Then
                            If Len(result) > 0 Then result = result & ", "
                            result = result & versionName
                        End If
                        Err.Clear
                    End If
                End If
            Next version
            GetFixVersions = result
        Else
            GetFixVersions = ""
        End If
    Else
        GetFixVersions = ""
    End If

    On Error GoTo 0
End Function

Private Function GetLabels(fields As Object) As String
    Dim labelsCollection As Object
    Dim label As Variant
    Dim result As String

    On Error Resume Next

    ' Get labels using .Item()
    Set labelsCollection = fields.Item("labels")

    If Err.Number = 0 And Not labelsCollection Is Nothing Then
        ' VBA-JSON returns Collection for arrays
        If TypeName(labelsCollection) = "Collection" Then
            result = ""
            For Each label In labelsCollection
                ' Labels are simple strings
                If Not IsEmpty(label) And Not IsNull(label) Then
                    If Len(result) > 0 Then result = result & ", "
                    result = result & CStr(label)
                End If
            Next label
            GetLabels = result
        Else
            GetLabels = ""
        End If
    Else
        GetLabels = ""
    End If

    On Error GoTo 0
End Function

Private Function GetEpicLink(fields As Object) As String
    Dim epicLink As String
    Dim fieldIds As Variant
    Dim fieldId As Variant
    Dim fieldValue As Variant
    Dim epicObj As Object

    ' Common Epic Link custom field IDs
    fieldIds = Array("customfield_10102", "customfield_10109", "customfield_10014", "customfield_10008", "customfield_10100", "customfield_10011")

    On Error Resume Next

    ' Try each possible field ID
    For Each fieldId In fieldIds
        Err.Clear

        ' Try to get the field value using .Item()
        fieldValue = fields.Item(CStr(fieldId))

        If Err.Number = 0 Then
            ' Check if it's a string (direct epic key)
            If VarType(fieldValue) = vbString Then
                If Len(fieldValue) > 0 Then
                    ' Validate that it looks like a Jira issue key (e.g., "PROJ-123")
                    ' Issue keys contain a hyphen and are at least 3 characters
                    If InStr(fieldValue, "-") > 0 And Len(fieldValue) >= 3 Then
                        Debug.Print "Epic Link found (string) in field " & fieldId & ": " & fieldValue
                        GetEpicLink = fieldValue
                        Exit Function
                    Else
                        Debug.Print "Ignoring invalid Epic Link format in field " & fieldId & ": " & fieldValue
                    End If
                End If
            ' Check if it's an object (epic link object with .key property)
            ElseIf IsObject(fieldValue) Then
                Set epicObj = fieldValue
                If Not epicObj Is Nothing Then
                    ' Try to get the key property
                    Err.Clear
                    epicLink = epicObj.Item("key")
                    If Err.Number = 0 And Len(epicLink) > 0 Then
                        ' Validate that it looks like a Jira issue key
                        If InStr(epicLink, "-") > 0 And Len(epicLink) >= 3 Then
                            Debug.Print "Epic Link found (object.key) in field " & fieldId & ": " & epicLink
                            GetEpicLink = epicLink
                            Exit Function
                        Else
                            Debug.Print "Ignoring invalid Epic Link format (object.key) in field " & fieldId & ": " & epicLink
                        End If
                    End If
                End If
            End If
        End If
    Next fieldId

    ' If not found in custom fields, try standard epic link field name
    Err.Clear
    fieldValue = fields.Item("epicLink")
    If Err.Number = 0 Then
        If VarType(fieldValue) = vbString And Len(fieldValue) > 0 Then
            ' Validate that it looks like a Jira issue key
            If InStr(fieldValue, "-") > 0 And Len(fieldValue) >= 3 Then
                GetEpicLink = fieldValue
                Exit Function
            End If
        ElseIf IsObject(fieldValue) Then
            Set epicObj = fieldValue
            If Not epicObj Is Nothing Then
                Err.Clear
                epicLink = epicObj.Item("key")
                If Err.Number = 0 And Len(epicLink) > 0 Then
                    ' Validate that it looks like a Jira issue key
                    If InStr(epicLink, "-") > 0 And Len(epicLink) >= 3 Then
                        GetEpicLink = epicLink
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    On Error GoTo 0

    ' Return empty string if not found
    GetEpicLink = ""
End Function

' ==========================================
' Function: GetEpicSummary
' Description: Get Epic Summary from Epic Link
' Parameters: epicKey - Epic issue key (e.g., "PROJ-123")
' Returns: Epic summary or empty string
' ==========================================
Private Function GetEpicSummary(epicKey As String) As String
    Dim http As Object
    Dim url As String
    Dim response As String
    Dim jsonResponse As Object
    Dim fieldsObj As Object
    Dim summary As String

    On Error Resume Next

    ' Return empty if no epic key
    If Len(epicKey) = 0 Then
        GetEpicSummary = ""
        Exit Function
    End If

    ' Build URL to fetch epic details
    url = JiraConfig.Config.JiraUrl & JiraConfig.GetApiPath() & "/issue/" & epicKey
    url = url & "?fields=summary"

    ' Create HTTP request
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then
        Set http = CreateObject("MSXML2.ServerXMLHTTP")
    End If
    If http Is Nothing Then
        Set http = CreateObject("MSXML2.XMLHTTP")
    End If

    If http Is Nothing Then
        GetEpicSummary = ""
        Exit Function
    End If

    ' Configure proxy if enabled
    If JiraConfig.Config.UseProxy Then
        Dim proxyUrl As String
        proxyUrl = JiraConfig.Config.ProxyServer & ":" & CStr(JiraConfig.Config.ProxyPort)
        http.setProxy 2, proxyUrl
        If Len(JiraConfig.Config.ProxyUsername) > 0 Then
            http.setProxyCredentials JiraConfig.Config.ProxyUsername, JiraConfig.Config.ProxyPassword
        End If
    End If

    ' Execute request
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", JiraConfig.GetAuthHeader()
    http.setRequestHeader "Accept", "application/json"
    http.Send

    ' Check response
    If http.Status = 200 Then
        response = http.responseText

        ' Parse JSON using VBA-JSON
        Set jsonResponse = JsonConverter.ParseJson(response)

        If Not jsonResponse Is Nothing Then
            ' Get fields object
            Set fieldsObj = jsonResponse.Item("fields")
            If Not fieldsObj Is Nothing Then
                ' Get summary
                summary = fieldsObj.Item("summary")
                If Err.Number = 0 And Len(summary) > 0 Then
                    GetEpicSummary = summary
                Else
                    GetEpicSummary = ""
                End If
            Else
                GetEpicSummary = ""
            End If
        Else
            GetEpicSummary = ""
        End If
    Else
        GetEpicSummary = ""
    End If

    Set http = Nothing
    Set jsonResponse = Nothing
    Set fieldsObj = Nothing

    On Error GoTo 0
End Function

' ==========================================
' Subroutine: EnsureSheetExists
' Description: Ensure sheet exists, create if not
' ==========================================
Private Sub EnsureSheetExists(sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
End Sub

' ==========================================
' Subroutine: CreateIssuesSheetLayout
' Description: Create layout for Issues sheet
' ==========================================
Private Sub CreateIssuesSheetLayout()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Issues")

    ' Clear sheet
    ws.Cells.Clear

    ' Headers
    ws.Range("A1").Value = "Key"
    ws.Range("B1").Value = "Summary"
    ws.Range("C1").Value = "Status"
    ws.Range("D1").Value = "Priority"
    ws.Range("E1").Value = "Assignee"
    ws.Range("F1").Value = "Epic Link"
    ws.Range("G1").Value = "Epic Summary"
    ws.Range("H1").Value = "Labels"
    ws.Range("I1").Value = "Fix Versions"
    ws.Range("J1").Value = "Sprint"
    ws.Range("K1").Value = "Created"
    ws.Range("L1").Value = "Total: 0"

    ' Format headers
    With ws.Range("A1:K1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Hide helper column (row number for detail view)
    ws.Columns("L:L").Hidden = True

    ' Freeze panes
    On Error Resume Next
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    ' Auto-filter (disable first if already enabled, for macOS compatibility)
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1:K1").AutoFilter
    On Error GoTo 0
End Sub

' ==========================================
' Subroutine: CreateFieldExplorerLayout
' Description: Create layout for FieldExplorer sheet
' ==========================================
Private Sub CreateFieldExplorerLayout()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("FieldExplorer")

    ' Clear sheet
    ws.Cells.Clear

    ' Title
    ws.Range("A1").Value = "Issue:"
    ws.Range("A1").Font.Bold = True
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 12

    ' Headers
    ws.Range("A2").Value = "Field Name"
    ws.Range("B2").Value = "Value"

    ' Format headers
    With ws.Range("A2:B2")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Column widths
    ws.Columns("A:A").ColumnWidth = 30
    ws.Columns("B:B").ColumnWidth = 60

    ' Freeze panes
    ws.Range("A3").Select
    ActiveWindow.FreezePanes = True
End Sub
