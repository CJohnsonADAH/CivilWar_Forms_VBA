VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditEngagements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If Me.Dirty Then
        Let Me!RevisedTimestamp.Locked = False
        Let Me!RevisedBy.Locked = False
        Let Me!RevisedTimestamp.Value = Now
        Let Me!RevisedBy.Value = getUserName
    End If
End Sub

Private Sub DoElaborateAlphanumSlug()
    If Len(Nz(Me!slug)) > 0 Then
        Dim sRegUnit As String
        Dim sUnitCo As String
        Dim cUnitCo As New Collection
        Dim sBit As String
        
        cUnitCo.Add Me!slug.Value
        
        Let sRegUnit = RegexReplace(Value:=Nz(Me!REGUNIT.Value), Pattern:="\s+Reg([.]|iment)?", Replace:="")
        Let sBit = RegexReplace(Value:=Trim(LCase(sRegUnit)), Pattern:="[^A-Za-z0-9]+", Replace:="-")
        If Len(sBit) > 0 Then cUnitCo.Add sBit
        
        Let sBit = RegexReplace(Value:=Trim(LCase(Nz(Me!BRANCH.Value))), Pattern:="[^A-Za-z0-9]+", Replace:="-")
        If Len(sBit) > 0 Then cUnitCo.Add sBit
        
        Let sBit = RegexReplace(Value:=Trim(LCase(Nz(Me!COMUNIT.Value))), Pattern:="[^A-Za-z0-9]+", Replace:="-")
        If Len(sBit) > 0 Then cUnitCo.Add "co-" & sBit
        
        Let sUnitCo = Join("-", List:=cUnitCo)
        Let Me!slug.Value = sUnitCo
        
        Set cUnitCo = Null
        
    End If
    
End Sub

Private Sub slug_DblClick(Cancel As Integer)
    If Len(Nz(Me!slug.Value)) = 0 Then
        DoFillInSlug
    Else
        DoElaborateAlphanumSlug
    End If
End Sub

Private Sub DoFillInSlug()
    Dim Control As Variant
    Dim Controls As New Collection
    Dim Elements As New Collection
    
    Controls.Add Item:=Me!EngagementName
    Controls.Add Item:=Me!EngagementState
    Controls.Add Item:=Me!EngagementLocation
        
    For Each Control In Controls
        Dim I As Long
        Dim ControlWords() As String
        Dim ControlText As String
        
        Let ControlWords = RegexSplit(Text:=Nz(Control.Value), Pattern:="[^A-Z0-9]+")
        For I = LBound(ControlWords) To UBound(ControlWords)
            Let ControlText = Trim(LCase(ControlWords(I)))
            If Len(ControlText) > 0 Then
                Elements.Add Item:=ControlText
            End If
        Next I
    Next Control
    
    Dim sSlug As String
    Let sSlug = Join(Delimiter:="-", List:=Elements)
    If Len(sSlug) > 0 Then
        Let Me!slug.Value = sSlug
    End If
    If Nz(Me!slug.Value) <> Nz(Me!slug.OldValue) Then
        HighlightField Field:=Me!slug
    End If
    
    Set Elements = Nothing
    Set Controls = Nothing
End Sub


Private Sub Form_Current()
    Let Me!RevisedTimestamp.Locked = True
    Let Me!RevisedBy.Locked = True

    If Len(Nz(Me!slug.Value)) = 0 Then
        DoFillInSlug
    End If

End Sub
