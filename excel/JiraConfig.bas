Attribute VB_Name = "JiraConfig"
Option Explicit

' ==========================================
' Module: JiraConfig
' Description: Configuration management for Jira connection
' Compatible with: Excel 2016+
' ==========================================

' API Version Enum
Public Enum ApiVersion
    SERVER_9_12_24 = 1      ' Jira Server 9.12.24 (API v2)
    CLOUD_CURRENT = 2       ' Jira Cloud (API v3)
End Enum

' Configuration structure
Public Type JiraConfiguration
    JiraUrl As String
    Username As String
    ApiToken As String
    MaxResults As Integer
    ApiVersionValue As ApiVersion
    ProxyServer As String
    ProxyPort As Integer
    ProxyUsername As String
    ProxyPassword As String
    UseProxy As Boolean
End Type

' Global configuration
Public Config As JiraConfiguration

' ==========================================
' Function: InitializeConfig
' Description: Initialize configuration with default values
' ==========================================
Public Sub InitializeConfig()
    Config.JiraUrl = ""
    Config.Username = ""
    Config.ApiToken = ""
    Config.MaxResults = 50
    Config.ApiVersionValue = CLOUD_CURRENT
    Config.ProxyServer = ""
    Config.ProxyPort = 8080
    Config.ProxyUsername = ""
    Config.ProxyPassword = ""
    Config.UseProxy = False
End Sub

' ==========================================
' Function: LoadConfigFromSheet
' Description: Load configuration from Config sheet
' ==========================================
Public Sub LoadConfigFromSheet()
    Dim ws As Worksheet
    Dim apiVersionStr As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Config")
    On Error GoTo 0

    If ws Is Nothing Then
        InitializeConfig
        Exit Sub
    End If

    ' Load configuration values
    Config.JiraUrl = Trim(ws.Range("B2").Value)

    ' Remove trailing slash from URL if present
    If Right(Config.JiraUrl, 1) = "/" Then
        Config.JiraUrl = Left(Config.JiraUrl, Len(Config.JiraUrl) - 1)
    End If

    Config.Username = ws.Range("B3").Value
    Config.ApiToken = ws.Range("B4").Value
    Config.MaxResults = ws.Range("B5").Value

    ' Parse API version
    apiVersionStr = ws.Range("B6").Value
    If apiVersionStr = "Jira Server 9.12.24" Then
        Config.ApiVersionValue = SERVER_9_12_24
    Else
        Config.ApiVersionValue = CLOUD_CURRENT
    End If

    ' Validate max results
    If Config.MaxResults < 1 Or Config.MaxResults > 1000 Then
        Config.MaxResults = 50
    End If

    ' Load proxy configuration
    Config.UseProxy = (ws.Range("B8").Value = "Yes" Or ws.Range("B8").Value = "Oui")
    Config.ProxyServer = ws.Range("B9").Value
    On Error Resume Next
    Config.ProxyPort = CInt(ws.Range("B10").Value)
    If Err.Number <> 0 Or Config.ProxyPort = 0 Then Config.ProxyPort = 8080
    On Error GoTo 0
    Config.ProxyUsername = ws.Range("B11").Value
    Config.ProxyPassword = ws.Range("B12").Value
End Sub

' ==========================================
' Function: SaveConfigToSheet
' Description: Save configuration to Config sheet
' ==========================================
Public Sub SaveConfigToSheet()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Config")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Config"
        Call CreateConfigSheetLayout(ws)
    End If

    ' Save configuration values
    ws.Range("B2").Value = Config.JiraUrl
    ws.Range("B3").Value = Config.Username
    ws.Range("B4").Value = Config.ApiToken
    ws.Range("B5").Value = Config.MaxResults

    ' Save API version
    If Config.ApiVersionValue = SERVER_9_12_24 Then
        ws.Range("B6").Value = "Jira Server 9.12.24"
    Else
        ws.Range("B6").Value = "Jira Cloud (Current)"
    End If

    ' Save proxy configuration
    ws.Range("B8").Value = IIf(Config.UseProxy, "Yes", "No")
    ws.Range("B9").Value = Config.ProxyServer
    ws.Range("B10").Value = Config.ProxyPort
    ws.Range("B11").Value = Config.ProxyUsername
    ws.Range("B12").Value = Config.ProxyPassword
End Sub

' ==========================================
' Function: CreateConfigSheetLayout
' Description: Create configuration sheet layout
' ==========================================
Private Sub CreateConfigSheetLayout(ws As Worksheet)
    ' Headers
    ws.Range("A1").Value = "Jira Configuration"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ' Configuration fields - Jira
    ws.Range("A2").Value = "Jira URL:"
    ws.Range("A3").Value = "Username (Email):"
    ws.Range("A4").Value = "API Token:"
    ws.Range("A5").Value = "Max Results:"
    ws.Range("A6").Value = "API Version:"

    ' Proxy section header
    ws.Range("A7").Value = ""
    ws.Range("A7").Font.Bold = True
    ws.Range("A7").Interior.Color = RGB(217, 217, 217)

    ' Configuration fields - Proxy
    ws.Range("A8").Value = "Use Proxy:"
    ws.Range("A9").Value = "Proxy Server:"
    ws.Range("A10").Value = "Proxy Port:"
    ws.Range("A11").Value = "Proxy Username:"
    ws.Range("A12").Value = "Proxy Password:"

    ' Format
    ws.Range("A2:A6").Font.Bold = True
    ws.Range("A8:A12").Font.Bold = True
    ws.Columns("A:A").ColumnWidth = 20
    ws.Columns("B:B").ColumnWidth = 40

    ' Add validation for API Version
    With ws.Range("B6").Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="Jira Server 9.12.24,Jira Cloud (Current)"
    End With

    ' Add validation for Use Proxy
    With ws.Range("B8").Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="Yes,No"
    End With

    ' Set default values
    ws.Range("B8").Value = "No"
    ws.Range("B10").Value = "8080"

    ' Add instructions
    ws.Range("A14").Value = "Instructions:"
    ws.Range("A14").Font.Bold = True
    ws.Range("A15").Value = "1. Enter your Jira URL WITHOUT trailing slash:"
    ws.Range("A16").Value = "   - Jira Cloud: https://your-domain.atlassian.net"
    ws.Range("A17").Value = "   - Jira Server: https://your-server.com/jira"
    ws.Range("A18").Value = "2. Select API Version based on your Jira instance"
    ws.Range("A19").Value = "3. Enter your Jira email address"
    ws.Range("A20").Value = "4. Generate and enter API token from:"
    ws.Range("A21").Value = "   https://id.atlassian.com/manage-profile/security/api-tokens"
    ws.Range("A22").Value = "5. Set maximum number of results (1-1000)"
    ws.Range("A23").Value = "6. (Optional) Configure proxy if required:"
    ws.Range("A24").Value = "   - Set Use Proxy to Yes"
    ws.Range("A25").Value = "   - Enter proxy server (e.g., proxy.company.com)"
    ws.Range("A26").Value = "   - Enter proxy port (default: 8080)"
    ws.Range("A27").Value = "   - Enter proxy credentials if required"
End Sub

' ==========================================
' Function: GetApiPath
' Description: Get API path based on version
' Returns: String - API path
' ==========================================
Public Function GetApiPath() As String
    If Config.ApiVersionValue = SERVER_9_12_24 Then
        GetApiPath = "/rest/api/2"
    Else
        GetApiPath = "/rest/api/3"
    End If
End Function

' ==========================================
' Function: GetSearchEndpoint
' Description: Get search endpoint based on API version
' Returns: String - Search endpoint path
' ==========================================
Public Function GetSearchEndpoint() As String
    If Config.ApiVersionValue = CLOUD_CURRENT Then
        GetSearchEndpoint = GetApiPath() & "/search/jql"
    Else
        GetSearchEndpoint = GetApiPath() & "/search"
    End If
End Function

' ==========================================
' Function: GetAuthHeader
' Description: Generate Basic Auth header
' Returns: String - Base64 encoded auth header
' ==========================================
Public Function GetAuthHeader() As String
    Dim credentials As String
    credentials = Config.Username & ":" & Config.ApiToken
'   GetAuthHeader = "Basic " & Base64Encode(credentials)
    GetAuthHeader = "Bearer " & Config.ApiToken 
End Function

' ==========================================
' Function: Base64Encode
' Description: Encode string to Base64
' Parameters: text - String to encode
' Returns: String - Base64 encoded string
' ==========================================
Private Function Base64Encode(ByVal text As String) As String
    Dim arrData() As Byte
    Dim objXML As Object
    Dim objNode As Object

    arrData = StrConv(text, vbFromUnicode)

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    Base64Encode = objNode.text

    Set objNode = Nothing
    Set objXML = Nothing
End Function

' ==========================================
' Function: IsConfigValid
' Description: Check if configuration is valid
' Returns: Boolean - True if valid, False otherwise
' ==========================================
Public Function IsConfigValid() As Boolean
    IsConfigValid = (Len(Config.JiraUrl) > 0 And _
                     Len(Config.Username) > 0 And _
                     Len(Config.ApiToken) > 0)
End Function
