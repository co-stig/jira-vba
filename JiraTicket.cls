VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JiraTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ticket As Object

Public Sub Init(ByRef ticket As Object)
    Set m_ticket = ticket
End Sub

Public Property Get status() As String
   status = m_ticket("fields")("status")("name")
End Property

Public Property Get assigneeName() As String
   assigneeName = m_ticket("fields")("assignee")("displayName")
End Property

Public Property Get name() As String
   name = m_ticket("fields")("summary")
End Property

Public Property Get updated() As String
   updated = Mid(m_ticket("fields")("updated"), 1, 10)
End Property

