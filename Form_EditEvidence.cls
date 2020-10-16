VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditEvidence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_AfterUpdate()

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    
    Dim oHistory As New cRevisedTable
    Dim Rs As DAO.Recordset
        
    If Me.Dirty Then
    
        If Me.NewRecord Then
        
            Let Me!PostedTimestamp.Locked = False
            Let Me!PostedBy.Locked = False
            Let Me!PostedTimestamp.Value = Now
            Let Me!PostedBy = oHistory.getUserName
            
        Else
        
            Let oHistory.IDName = "ID"
    
            Set Rs = Me.RecordsetClone
            oHistory.DoStoreRevision ID:=Me!ID.Value, Rs:=Rs
            Set Rs = Nothing
            
            Let Me!RevisedTimestamp.Locked = False
            Let Me!RevisedBy.Locked = False
            Let Me!RevisedTimestamp.Value = Now
            Let Me!RevisedBy.Value = oHistory.getUserName

        End If
        
    End If
    
    Set oHistory = Nothing
    
    Exit Sub

End Sub
