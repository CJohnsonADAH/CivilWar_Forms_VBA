VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRevisedTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private sIDName As String
Private sHistoryTable As String

Public Property Get IDName() As String
    Let IDName = sIDName
End Property

Public Property Let IDName(ByVal RHS As String)
    Let sIDName = RHS
End Property

Public Property Get HistoryTable() As String
    Let HistoryTable = sHistoryTable
End Property

Public Property Let HistoryTable(ByVal RHS As String)
    Let sHistoryTable = RHS
End Property

Public Sub DoStoreRevision(ByVal ID As Variant, ByRef Rs As DAO.Recordset, Optional ByVal User As String, Optional ByVal LatestValue As Boolean)
    Dim revRs As DAO.Recordset
    Dim Table As String
    Dim Field As Variant
    Dim sqlID As String
    
    If Len(IDName) = 0 Then
        Let IDName = "ID"
    End If
    
    If TypeName(ID) = "String" Then
        Let sqlID = "'" & Replace(ID, "'", "''") & "'"
    Else
        Let sqlID = ID
    End If
    
    Rs.FindFirst "[" & IDName & "]=" & sqlID
    Let Table = Rs.Fields(IDName).SourceTable
    If Len(HistoryTable) = 0 Then
        Let HistoryTable = Table & "RevisionHistory"
    End If
    
    If Not Rs.EOF Then
        Set revRs = CurrentDb.OpenRecordset(HistoryTable)
        revRs.AddNew
        For Each Field In Rs.Fields
            On Error Resume Next
            If LatestValue Then
                Let revRs.Fields(Field.Name) = Rs.Fields(Field.Name).Value
            Else
                Let revRs.Fields(Field.Name) = Field.Value
            End If
            On Error GoTo 0
        Next Field
        On Error Resume Next
        revRs.Fields("RevTimestamp").Value = Now
        revRs.Fields("RevUserName").Value = IIf(Len(User) = 0, getUserName, User)
        On Error GoTo 0
        revRs.Update
            
        revRs.Close
    End If

End Sub


'**
'* getUserName: get a username representing the current user of the system or application
'*
'* @return String a username representing the current user of the system or application
'**
Public Function getUserName() As String
    Dim sUserName As String
    
    Let sUserName = Environ$("Username")
    If Len(sUserName) = 0 Then
        Let sUserName = Application.CurrentUser
    End If
    
    Let getUserName = sUserName
End Function
