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
                ws.Cells(row, 1).Value = GetDictValue(issue, "key")
                ws.Cells(row, 2).Value = GetDictValue(fields, "summary")
                ws.Cells(row, 3).Value = GetDictNestedValue(fields, "status", "name")
                ws.Cells(row, 4).Value = GetDictNestedValue(fields, "priority", "name")
                ws.Cells(row, 5).Value = GetDictNestedValue(fields, "assignee", "displayName")
                ws.Cells(row, 6).Value = GetEpicLink(fields)
                ws.Cells(row, 7).Value = GetLabels(fields)
                ws.Cells(row, 8).Value = GetFixVersions(fields)
                Dim sprintValue As String
                sprintValue = GetSprints(fields)
                Debug.Print "Row " & row & " Sprint value: '" & sprintValue & "'"
                ws.Cells(row, 9).Value = sprintValue
                ws.Cells(row, 10).Value = GetDictValue(fields, "created")
            End If

            ' Store full issue data in hidden column for detail view
            ws.Cells(row, 11).Value = row
            ws.Cells(row, 11).NumberFormat = "0"

            row = row + 1
        End If
    Next issue

    ' Auto-fit columns
    ws.Columns("A:J").AutoFit

    ' Update result count
    ws.Range("K1").Value = "Total: " & issues.Count
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
    Dim scriptControl As Object
    Dim commonFields As Variant
    Dim i As Long
    Dim fieldKey As String
    Dim fieldName As String
    Dim fieldValue As String

    ' Clear existing data
    ws.Rows("3:" & ws.Rows.Count).ClearContents

    ' Set issue key
    ws.Range("B1").Value = issueKey

    ' Create ScriptControl for accessing values
    Set scriptControl = CreateObject("ScriptControl")
    scriptControl.Language = "JScript"

    ' Parse JSON and extract issue
    ' For Jira Server, the response has issues array; for Cloud, it's the issue directly
    Dim jsCode As String
    jsCode = "var rawData = " & issueJson & ";"
    jsCode = jsCode & "var issueObj = rawData.issues ? rawData.issues[0] : rawData;"
    scriptControl.AddCode jsCode

    Debug.Print "JSON loaded into ScriptControl"

    ' Define common Jira fields to display
    commonFields = Array("summary", "description", "status", "priority", "assignee", _
                         "reporter", "created", "updated", "resolutiondate", "duedate", _
                         "issuetype", "project", "components", "labels", "fixVersions", _
                         "versions", "timeoriginalestimate", "timeestimate", "timespent", _
                         "aggregatetimeoriginalestimate", "aggregatetimeestimate", "aggregatetimespent", _
                         "resolution", "customfield_10109", "customfield_10014", "customfield_10008", "customfield_10011")

    ' Add function to get field value - build in parts to avoid line continuation limit
    jsCode = "function getFieldVal(fieldKey) { try { var fields = issueObj.fields;"
    jsCode = jsCode & "if (!fields) return '[No fields]';"
    jsCode = jsCode & "if (!fields.hasOwnProperty(fieldKey)) return '';"
    jsCode = jsCode & "var val = fields[fieldKey];"
    jsCode = jsCode & "if (val === null || val === undefined) return '';"
    jsCode = jsCode & "if (typeof val === 'string') return val;"
    jsCode = jsCode & "if (typeof val === 'number') return String(val);"
    jsCode = jsCode & "if (typeof val === 'boolean') return val ? 'true' : 'false';"
    jsCode = jsCode & "if (typeof val === 'object') {"
    jsCode = jsCode & "if (val.constructor && val.constructor.toString().indexOf('Array') > -1) {"
    jsCode = jsCode & "var items = []; for (var i = 0; i < val.length; i++) {"
    jsCode = jsCode & "var item = val[i];"
    jsCode = jsCode & "if (typeof item === 'string') items.push(item);"
    jsCode = jsCode & "else if (item && item.name) items.push(item.name);"
    jsCode = jsCode & "else if (item && item.value) items.push(item.value);"
    jsCode = jsCode & "else if (item && item.key) items.push(item.key); }"
    jsCode = jsCode & "return items.length > 0 ? items.join(', ') : ''; }"
    jsCode = jsCode & "if (val.name) return val.name;"
    jsCode = jsCode & "if (val.value) return val.value;"
    jsCode = jsCode & "if (val.displayName) return val.displayName;"
    jsCode = jsCode & "if (val.key) return val.key;"
    jsCode = jsCode & "if (val.self) return ''; return ''; }"
    jsCode = jsCode & "return String(val);"
    jsCode = jsCode & "} catch(e) { return ''; } }"
    scriptControl.AddCode jsCode

    ' Add function to get Epic Link
    jsCode = "function getEpicLink() { try { var fields = issueObj.fields;"
    jsCode = jsCode & "if (!fields) return '';"
    jsCode = jsCode & "var epicFieldIds = ['customfield_10109','customfield_10014','customfield_10008','customfield_10100','customfield_10011'];"
    jsCode = jsCode & "for (var i = 0; i < epicFieldIds.length; i++) {"
    jsCode = jsCode & "var val = fields[epicFieldIds[i]];"
    jsCode = jsCode & "if (val !== null && val !== undefined && val !== '') {"
    jsCode = jsCode & "if (typeof val === 'string') return val;"
    jsCode = jsCode & "if (typeof val === 'object' && val.key) return val.key; } }"
    jsCode = jsCode & "return ''; } catch(e) { return ''; } }"
    scriptControl.AddCode jsCode

    ' Add debug function to check field type
    jsCode = "function getFieldType(fieldKey) { try {"
    jsCode = jsCode & "var fields = issueObj.fields;"
    jsCode = jsCode & "if (!fields || !fields.hasOwnProperty(fieldKey)) return 'missing';"
    jsCode = jsCode & "var val = fields[fieldKey];"
    jsCode = jsCode & "if (val === null) return 'null';"
    jsCode = jsCode & "if (val === undefined) return 'undefined';"
    jsCode = jsCode & "return typeof val;"
    jsCode = jsCode & "} catch(e) { return 'error: ' + e.message; } }"
    scriptControl.AddCode jsCode

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

        ' Check field type for debugging
        Dim fieldType As String
        On Error Resume Next
        fieldType = scriptControl.Run("getFieldType", fieldKey)
        If Err.Number <> 0 Then
            fieldType = "error"
            Err.Clear
        End If

        Debug.Print "Field '" & fieldKey & "' type: " & fieldType

        ' Get field value
        fieldValue = scriptControl.Run("getFieldVal", fieldKey)
        If Err.Number <> 0 Then
            Debug.Print "Error getting field '" & fieldKey & "': " & Err.Description
            fieldValue = ""
            Err.Clear
        End If
        On Error GoTo 0

        ' Only display if field has a value
        If Len(fieldValue) > 0 Then
            Debug.Print "Field '" & fieldKey & "' = '" & fieldValue & "'"
            ws.Cells(row, 1).Value = fieldName
            ws.Cells(row, 2).Value = fieldValue
            row = row + 1
        End If
    Next i

    ' Add Epic Link field (try multiple custom field IDs)
    On Error Resume Next
    fieldValue = scriptControl.Run("getEpicLink")
    If Err.Number = 0 And Len(fieldValue) > 0 Then
        ws.Cells(row, 1).Value = "Epic Link"
        ws.Cells(row, 2).Value = fieldValue
        row = row + 1
    End If
    Err.Clear
    On Error GoTo 0

    ' Auto-fit columns
    ws.Columns("A:B").AutoFit

    Set scriptControl = Nothing
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
    ws.Range("G1").Value = "Labels"
    ws.Range("H1").Value = "Fix Versions"
    ws.Range("I1").Value = "Sprint"
    ws.Range("J1").Value = "Created"
    ws.Range("K1").Value = "Total: 0"

    ' Format headers
    With ws.Range("A1:J1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Hide helper column (row number for detail view)
    ws.Columns("K:K").Hidden = True

    ' Freeze panes
    On Error Resume Next
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    ' Auto-filter (disable first if already enabled, for macOS compatibility)
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1:J1").AutoFilter
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
