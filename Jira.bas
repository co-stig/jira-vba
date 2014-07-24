Attribute VB_Name = "Jira"
Public Function Connect(url As String, user As String, password As String) As JiraServer
    Set Connect = New JiraServer
    Connect.InitializeProperties url:=url, user:=user, password:=password
End Function
