VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JiraServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_url As String
Private m_user As String
Private m_password As String
Private m_auth As String

Private Function HttpRequest(ByVal method As String, ByVal relativeUrl As String, ByVal data As String, ByVal auth As Boolean, ByVal responseRequired As Boolean) As Object
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    Let url = m_url + relativeUrl

    request.Open method, url, False
    request.setRequestHeader "Content-Type", "application/json"
    request.setRequestHeader "Accept", "application/json"
    
    If auth Then
        If m_auth = "" Then
            Call HttpAuth
        End If
        request.setRequestHeader "Cookie", m_auth
    End If
    
    request.send data
    Debug.Print "HttpRequest to '" & url & "': " & request.responseText
    
    Dim lib As New JSONLib
    If responseRequired Then
        Set HttpRequest = lib.parse(request.responseText)
    End If
    
    ' TODO: Handle "not authenticated" errors here by logging in again
End Function

Private Sub HttpAuth()
    Dim resp As Object
    Set resp = HttpRequest("POST", "/rest/auth/latest/session", " {""username"" : """ & m_user & """, ""password"" : """ & m_password & """}""", False, True)
    m_auth = "JSESSIONID=" & resp("session")("value") & "; Path=/"
End Sub

Public Function FindTicket(ByVal id As String) As JiraTicket
    Set FindTicket = New JiraTicket
    FindTicket.Init HttpRequest("GET", "/rest/api/latest/issue/" + id, "", True, True)
End Function

Public Function CreateTicket(ByVal name As String, ByVal text As String, ByVal objectType As String) As String
    Let typeId = "3"
    If objectType = "D" Then
        typeId = "10000"
    End If
    Let text = "{""fields"": {""project"": {""id"": ""10000""}, ""summary"": """ & name & """, ""issuetype"": {""id"": """ & typeId & """}, ""assignee"": {""name"": ""admin""}, ""reporter"": {""name"": ""admin""}, ""priority"": {""id"": ""3""}, ""description"": """ + text + """}}"
    Dim resp As Object
    Set resp = HttpRequest("POST", "/rest/api/latest/issue", text, True, True)
    CreateTicket = resp("key")
End Function

' linkType: Blocks / Produces
Public Sub CreateLink(ByVal fromIssue As String, ByVal toIssue As String, ByVal linkType As String)
    Debug.Print "CreateLink '" & fromIssue & "' -> '" & toIssue & "' as '" & linkType & "'"
    Let text = "{""type"": {""name"": """ & linkType & """}, ""inwardIssue"": {""key"": """ & fromIssue & """}, ""outwardIssue"": {""key"": """ & toIssue & """}}"
    Call HttpRequest("POST", "/rest/api/2/issueLink", text, True, False)
End Sub

Public Sub CreateRemoteLink(ByVal issue As String, ByVal url As String, ByVal description As String)
    Debug.Print "CreateRemoteLink '" & issue & "' -> '" & url & "' as '" & description & "'"
    Let text = "{""object"":{""url"":""" & url & """,""title"":""" & description & """}}"
    Call HttpRequest("POST", "/rest/api/2/issue/" & issue & "/remotelink", text, True, False)
End Sub

Public Sub InitializeProperties(url As String, user As String, password As String)
    m_url = url
    m_user = user
    m_password = password
End Sub
