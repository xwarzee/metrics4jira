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
            ' Get fields object using CallByName for JScript compatibility
            On Error Resume Next
            Set fields = CallByName(issue, "fields", VbGet)
            On Error GoTo 0

            If Not fields Is Nothing Then
                ws.Cells(row, 1).Value = GetValue(issue, "key")
                ws.Cells(row, 2).Value = GetValue(fields, "summary")
                ws.Cells(row, 3).Value = GetNestedValue(fields, "status", "name")
                ws.Cells(row, 4).Value = GetNestedValue(fields, "priority", "name")
                ws.Cells(row, 5).Value = GetNestedValue(fields, "assignee", "displayName")
                ws.Cells(row, 6).Value = GetEpicLink(fields)
                ws.Cells(row, 7).Value = GetValue(fields, "created")
            End If

            ' Store full issue data in hidden column for detail view
            ws.Cells(row, 8).Value = row
            ws.Cells(row, 8).NumberFormat = "0"

            row = row + 1
        End If
    Next issue

    ' Auto-fit columns
    ws.Columns("A:G").AutoFit

    ' Update result count
    ws.Range("H1").Value = "Total: " & issues.Count
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

    ' Fetch issue details
    jql = "key = " & issueKey
    Set issues = JiraApiClient.SearchIssues(jql, 0, 1)

    If issues.Count > 0 Then
        Set issue = issues(1)
        Set fields = issue("fields")

        ' Display field explorer
        Call DisplayFieldExplorer(wsExplorer, issueKey, fields)
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

    ' Clear existing data
    ws.Rows("3:" & ws.Rows.Count).ClearContents

    ' Set issue key
    ws.Range("B1").Value = issueKey

    ' Display fields
    row = 3
    For Each key In GetObjectKeys(fields)
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

        ' Get field value
        fieldValue = FormatFieldValue(fields(key))

        ws.Cells(row, 1).Value = fieldName
        ws.Cells(row, 2).Value = fieldValue

        row = row + 1
    Next key

    ' Auto-fit columns
    ws.Columns("A:B").AutoFit
End Sub

' ==========================================
' Function: GetObjectKeys
' Description: Get keys from object (JScript object)
' Parameters: obj - Object to get keys from
' Returns: Array of keys
' ==========================================
Private Function GetObjectKeys(obj As Object) As Variant
    Dim scriptControl As Object
    Dim keys As Object

    Set scriptControl = CreateObject("ScriptControl")
    scriptControl.Language = "JScript"
    scriptControl.AddCode "function getKeys(obj) { var arr = []; for (var k in obj) { arr.push(k); } return arr; }"

    Set keys = scriptControl.Run("getKeys", obj)
    GetObjectKeys = keys.toArray()

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
' Function: GetEpicLink
' Description: Get Epic Link from fields (tries multiple custom field IDs)
' ==========================================
Private Function GetEpicLink(fields As Object) As String
    Dim epicLink As String
    Dim fieldIds As Variant
    Dim fieldId As Variant

    ' Common Epic Link custom field IDs
    ' customfield_10014 - Jira Cloud default
    ' customfield_10008 - Common in Jira Server
    ' customfield_10100 - Another common ID
    fieldIds = Array("customfield_10014", "customfield_10008", "customfield_10100", "customfield_10011")

    On Error Resume Next

    ' Try each possible field ID
    For Each fieldId In fieldIds
        Err.Clear
        epicLink = GetValue(fields, CStr(fieldId))

        If Err.Number = 0 And Len(epicLink) > 0 Then
            GetEpicLink = epicLink
            Exit Function
        End If
    Next fieldId

    ' If not found in custom fields, try standard epic link field name
    Err.Clear
    epicLink = GetValue(fields, "epicLink")
    If Err.Number = 0 And Len(epicLink) > 0 Then
        GetEpicLink = epicLink
        Exit Function
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
    ws.Range("G1").Value = "Created"
    ws.Range("I1").Value = "Total: 0"

    ' Format headers
    With ws.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Hide helper column (row number for detail view)
    ws.Columns("H:H").Hidden = True

    ' Freeze panes
    On Error Resume Next
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    ' Auto-filter (disable first if already enabled, for macOS compatibility)
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1:G1").AutoFilter
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
