VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditSoldierRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Function EngagementsBreaker() As String
    Let EngagementsBreaker = "^(Killed|Ab|C(apt|[.])|Pr(es|[.:\s])|Sev)"
End Function

Private Sub AUTH_DblClick(Cancel As Integer)
    Dim sAuth As String
    Dim sAuthRegex As String
        
    Let sAuth = Me!AUTH.Value
    Let sAuthRegex = "Authority cont[.]|Auth(or(ity)?)?(\s+con[']t)?[:.]?\s*"
    
    DoConsiderForJoining Text:=sAuth, Field:=Me!ENGAGE2, PrefixPattern:=sAuthRegex
    DoConsiderForJoining Text:=sAuth, Field:=Me!REMARKS, PrefixPattern:=sAuthRegex
    DoConsiderForJoining Text:=sAuth, Field:=Me!REMARKS2, PrefixPattern:=sAuthRegex
    
    If sAuth <> Me!AUTH.Value Then
        Let Me!AUTH.Value = sAuth
        HighlightField Field:=Me!AUTH
    End If
    
End Sub

Private Sub cmdEngage2ToRemarks_Click()
    Dim sRemarks As String
    Dim sEngage2 As String
    
    If Len(Nz(Me!ENGAGE2.Value)) > 0 Then
        Let sEngage2 = Nz(Me!ENGAGE2.Value)
        Let sEngage2 = RegexReplace(sEngage2, "^Remarks( con[']t)?[:]\s*", "")
        
        Let sRemarks = FieldsJoined(sEngage2, Nz(Me!REMARKS.Value))
        Let Me!REMARKS.Value = sRemarks
        Let Me!ENGAGE2.Value = Null
        
        DoFilterREMARKSField Field:=Me!REMARKS
        HighlightField Field:=Me!ENGAGE2
    End If
End Sub

Private Sub cmdEngageDates_Click()
    Let Me.Filter = "Len(Nz([ENGAGE2])) > 0 And ([ENGAGE] Like '*##/*' OR [ENGAGE] Like '*Pres*')"
    Let Me.FilterOn = True

End Sub

Private Sub DoPrepEngageToRemarks()
    Dim sEngage As String
    Dim sEngage2 As String
    Dim vEngage As Variant
    
    Let vEngage = Me!ENGAGE.Value
    Let sEngage = Nz(Me!ENGAGE.Value)
    
    Let sEngage = RegexReplace(Me!ENGAGE.Value, "^Remarks:\s+", "")
    If Len(Nz(Me!ENGAGE2.Value)) > 0 Then
        Let sEngage2 = Nz(Me!ENGAGE2.Value)
        Let sEngage2 = RegexReplace(sEngage2, "^Remarks( con[']t)?[:]\s*", "")
            
        Let sEngage = FieldsJoined(sEngage, sEngage2)
        Let Me!ENGAGE2.Value = Null
        HighlightField Field:=Me!ENGAGE2
    End If
        
    Let Me!REMARKS.Value = RegexReplace(Me!REMARKS.Value, "^(Con[']t)[:]", "")
    
    Let Me!ENGAGE.Value = sEngage
    
End Sub

Private Sub cmdEngageToRemarks_Click()
    DoPrepEngageToRemarks
    Let Me!REMARKS.Value = FieldsJoined(Me!ENGAGE.Value, Me!REMARKS.Value)
    DoFilterREMARKSField Field:=Me!REMARKS
    
    Let Me!ENGAGE.Value = Null
    HighlightField Field:=Me!ENGAGE

End Sub

Private Sub cmdEngageToRemarksBack_Click()
    DoPrepEngageToRemarks
    Let Me!REMARKS.Value = FieldsJoined(Me!REMARKS.Value, Me!ENGAGE.Value)
    DoFilterREMARKSField Field:=Me!REMARKS
    
    Let Me!ENGAGE.Value = Null
    HighlightField Field:=Me!ENGAGE
    
End Sub

Private Sub cmdJoinEm_Click()
    DoProcessENGAGE_Field
    DoProcessREMARKS_Field
End Sub

Private Sub cmdNextRecord_Click()
    DoCmd.GoToRecord Record:=acNext
End Sub

Private Sub cmdRemarksToEngage_Click()
    Dim sRemarks As String
    Dim sEngage As String
    Dim vRemainder As Variant
    Dim vRemainderFragment As Variant
    Dim reRemarksProperties As String
    
    Let sEngage = Nz(Me!ENGAGE.Value)
    If Len(Nz(Me!REMARKS.Value)) > 0 Then
        Let sRemarks = Nz(Me!REMARKS.Value)
        Let sRemarks = RegexReplace(sRemarks, "^Engage( con[']t)?[:]\s*", "")
        Let vRemainder = Null
        
        Let reRemarksProperties = "^(.*)(\b(Age|Rank|Residence|Res[.]|P[.]?\s*O([.]\s*|[:]|\s+))[:,]?\s+.*)$"
        Do While RegexMatch(sRemarks, reRemarksProperties)
            Let vRemainderFragment = RegexReplace(sRemarks, reRemarksProperties, "$2")
            If IsNull(vRemainder) Then
                Let vRemainder = ""
            End If
            Let vRemainder = vRemainderFragment & vRemainder
            Let sRemarks = RegexReplace(sRemarks, reRemarksProperties, "$1")
        Loop
        
        Let sEngage = FieldsJoined(sEngage, sRemarks)
        Let Me!ENGAGE.Value = sEngage
        Let Me!REMARKS.Value = vRemainder
        HighlightField Field:=Me!ENGAGE
        HighlightField Field:=Me!REMARKS
    End If

End Sub

Private Sub ENGAGE_DblClick(Cancel As Integer)
    DoProcessENGAGE_Field
End Sub

Private Sub DoProcessENGAGE_Field()
    Dim sEngage As String
    Dim sEngage2 As String
    Dim vEngage As Variant
    
    Let vEngage = Me!ENGAGE.Value
    Let sEngage = Nz(Me!ENGAGE.Value)
    
    Dim sRemarks As String
    If RegexMatch(Value:=Me!ENGAGE.Value, Pattern:="^Remarks:\s+") Then
        Let sRemarks = RegexReplace(Me!ENGAGE.Value, "^Remarks:\s+", "")
        If Len(Nz(Me!ENGAGE2.Value)) > 0 Then
            Let sEngage2 = Nz(Me!ENGAGE2.Value)
            Let sEngage2 = RegexReplace(sEngage2, "^Remarks( con[']t)?[:]\s*", "")
            
            Let sRemarks = FieldsJoined(sRemarks, sEngage2)
            Let Me!ENGAGE2.Value = Null
            HighlightField Field:=Me!ENGAGE2
        End If
        
        Let Me!REMARKS.Value = RegexReplace(Me!REMARKS.Value, "^(Con[']t)[:]", "")
        
        Let Me!REMARKS.Value = FieldsJoined(sRemarks, Me!REMARKS.Value)
        Let Me!ENGAGE.Value = Null
        HighlightField Field:=Me!ENGAGE
    Else
        If Len(Nz(Me!ENGAGE2.Value)) > 0 Then
            DoJoinEngagementFields RunningText:=sEngage, ENGAGE:=Me!ENGAGE, ENGAGE2:=Me!ENGAGE2
        End If
    
        DoConsiderForJoining Text:=sEngage, Field:=Me!REMARKS, PrefixPattern:=EngagementsContdCombinedRegex, Breaker:=EngagementsBreaker
        DoConsiderForJoining Text:=sEngage, Field:=Me!REMARKS2, PrefixPattern:=EngagementsContdCombinedRegex, Breaker:=EngagementsBreaker
    
        Let sEngage = RegexReplace(Value:=sEngage, Pattern:="([;:,.])\s*", Replace:="$1 ")
    
        Let Me!ENGAGE.Value = sEngage
    
        If (CStr(Nz(vEngage)) <> sEngage) Then
            HighlightField Field:=Me!ENGAGE
        End If
    End If

End Sub

Private Sub ENGAGE_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Me!ENGAGE.Tag = "Modified") And (Shift And 1) And (Button = 1) Then
        On Error GoTo RecordSave_Bounced:
        Let Me.Dirty = False
        On Error GoTo 0
        
        DoCmd.GoToRecord Record:=acNext
    End If
    
    If (Shift And 2) And (Button = 1) Then
        If Len(Nz(Me!REMARKS.Value)) > 0 Then
            Let Me!ENGAGE.Value = FieldsJoined(Text1:=Me!ENGAGE.Value, Text2:=Me!REMARKS.Value)
            Let Me!REMARKS.Value = ""
            HighlightField Field:=Me!REMARKS
        ElseIf Len(Nz(Me!REMARKS2.Value)) > 0 Then
            Let Me!ENGAGE.Value = FieldsJoined(Text1:=Me!ENGAGE.Value, Text2:=Me!REMARKS2.Value)
            Let Me!REMARKS2.Value = ""
            HighlightField Field:=Me!REMARKS2
        End If
    End If
    Exit Sub
    
RecordSave_Bounced:
    DoElaborateAlphanumSlug
    Resume
    
End Sub

Private Sub DoElaborateAlphanumSlug()
    If Len(Nz(Me!slug)) > 0 Then
        Dim sRegUnit As String
        Dim sUnitCo As String
        Dim iLenRegBranchCo As Integer
        
        Let sRegUnit = RegexReplace(Value:=Nz(Me!REGUNIT.Value), Pattern:="\s+Reg([.]|iment)?", Replace:="")
        Let iLenRegBranchCo = Len(Nz(Me!REGUNIT.Value))
        
        Let sUnitCo = RegexReplace(Value:=Trim(LCase(sRegUnit)), Pattern:="[^A-Za-z0-9]+", Replace:="-")
        Let sUnitCo = sUnitCo & "-" & RegexReplace(Value:=Trim(LCase(Nz(Me!BRANCH.Value))), Pattern:="[^A-Za-z0-9]+", Replace:="-")
        Let iLenRegBranchCo = iLenRegBranchCo + Len(Nz(Me!BRANCH.Value))
        
        If Len(Nz(Me!COMUNIT.Value)) > 0 Then
            Let sUnitCo = sUnitCo & "-co-" & RegexReplace(Value:=Trim(LCase(Nz(Me!COMUNIT.Value))), Pattern:="[^A-Za-z0-9]+", Replace:="-")
            Let iLenRegBranchCo = iLenRegBranchCo + Len(Nz(Me!COMUNIT.Value))
        End If
        
        If iLenRegBranchCo > 0 Then
            Let Me!slug.Value = Me!slug.Value & "-" & sUnitCo
        Else
            Let Me!slug.Value = Me!slug.Value & "-" & VBA.FORMAT$(Me!recnumber.Value, "0")
        End If
        
    End If
    
End Sub

Private Sub ENGAGE2_DblClick(Cancel As Integer)
    Dim sRemarks As String
    If RegexMatch(Value:=Me!ENGAGE2.Value, Pattern:="^Remarks:\s+") Then
        Let sRemarks = RegexReplace(Me!ENGAGE2.Value, "^Remarks:\s+", "")
        Let Me!REMARKS.Value = RegexReplace(Me!REMARKS.Value, "^(Con[']t)[:]", "")
        
        Let Me!REMARKS.Value = FieldsJoined(sRemarks, Me!REMARKS.Value)
        Let Me!ENGAGE2.Value = Null
        HighlightField Field:=Me!ENGAGE2
    End If
    
End Sub

Private Sub ENGAGE2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Me!ENGAGE2.Tag = "Modified") And (Shift And 1) And (Button = 1) Then
        Let Me.Dirty = False
        DoCmd.GoToRecord Record:=acNext
    End If

End Sub

Private Sub Form_AfterUpdate()
    DoResetHighlightedFields
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Dim oHistory As cRevisedTable
    Dim Rs As DAO.Recordset
        
    If Me.Dirty Then
        Set oHistory = New cRevisedTable
        Let oHistory.IDName = "recnumber"

        Set Rs = Me.RecordsetClone
        oHistory.DoStoreRevision ID:=Me!recnumber.Value, Rs:=Rs
        Set Rs = Nothing
        
        Let Me!RevisedTimestamp.Locked = False
        Let Me!RevisedBy.Locked = False
        Let Me!RevisedTimestamp.Value = Now
        Let Me!RevisedBy.Value = oHistory.getUserName
              
        Dim Dupe As Variant
        Let Dupe = DLookup(Expr:="recnumber", Domain:="Cw", Criteria:="[slug]='" & Nz(Me!slug.Value) & "'")
        If Not IsNull(Dupe) Then
            If Dupe <> Me!recnumber.Value Then
                DoElaborateAlphanumSlug
            End If
        End If
    End If
    Exit Sub
    
End Sub

Private Sub Form_Current()
    Let Me!RevisedTimestamp.Locked = True
    Let Me!RevisedBy.Locked = True
    
    Let Me!ENGAGE2.Visible = (Len(Nz(Me!ENGAGE2.Value)) > 0)
    Let Me!cmdEngage2ToRemarks.Visible = Me!ENGAGE2.Visible
    Let Me!ENGAGE2.Height = IIf(Me!ENGAGE2.Visible, 810, 10)
    Let Me!cmdEngage2ToRemarks.Height = IIf(Me!cmdEngage2ToRemarks.Visible, 810, 10)
    Let Me!REMARKS2.Visible = (Len(Nz(Me!REMARKS2.Value)) > 0)
    Let Me!REMARKS2.Height = IIf(Me!REMARKS2.Visible, 810, 10)

    DoResetHighlightedFields
    If Len(Nz(Me!slug.Value)) = 0 Then
        DoFillInSlug
    End If

End Sub

Private Sub Label30_DblClick(Cancel As Integer)
    Let Me.Filter = "Len(Nz([ENGAGE2])) > 0"
    Let Me.FilterOn = Not Me.FilterOn
End Sub

Private Sub lblRemarks_DblClick(Cancel As Integer)
    If Me.Filter <> "Len(Nz([REMARKS2])) > 0" Then
        Let Me.Filter = "Len(Nz([REMARKS2])) > 0"
        Let Me.FilterOn = False
    End If
    Let Me.FilterOn = Not Me.FilterOn
End Sub

Private Sub REMARKS_DblClick(Cancel As Integer)
    DoProcessREMARKS_Field
End Sub

Private Sub DoProcessREMARKS_Field()
    Dim sEngage2 As String
    Dim sRemarksRegex As String
    
    Dim sRemarks As String
    Let sRemarks = Nz(Me!REMARKS.Value)
    
    If Len(Nz(Me!REMARKS2.Value)) > 0 Then
        Let Me!REMARKS.Value = FieldsJoined(Nz(Me!REMARKS.Value), Nz(Me!REMARKS2.Value))
        Let Me!REMARKS2.Value = ""
        HighlightField Field:=Me!REMARKS2
    End If
    
    If Len(Nz(Me!ENGAGE2.Value)) > 0 Then
        Let sRemarksRegex = "^Rem(arks?)?(\s+con[']t)?[:]?\s+"
        If RegexMatch(Value:=Me!ENGAGE2.Value, Pattern:=sRemarksRegex) Then
            Let sEngage2 = Me!ENGAGE2.Value
            Let sEngage2 = RegexReplace(sEngage2, sRemarksRegex, "")
            Let Me!REMARKS.Value = FieldsJoined(Me!REMARKS.Value, sEngage2)
            Let Me!ENGAGE2.Value = ""
            HighlightField Field:=Me!ENGAGE2
        End If
    End If
    
    DoFilterREMARKSField Field:=Me!REMARKS
    
    If sRemarks <> Nz(Me!REMARKS.Value) Then
        HighlightField Field:=Me!REMARKS
    End If
    
End Sub

Private Sub REMARKS2_BeforeUpdate(Cancel As Integer)

End Sub

Private Sub REMARKS2_DblClick(Cancel As Integer)
    Dim sOTHERINFO As String
    Dim sOTHERINFO_Original As String
    
    Dim sAUTHORITY As String
    Dim sAUTHORITY_Original As String
    
    Let sOTHERINFO = Nz(Me!OTHERINFO.Value)
    Let sOTHERINFO_Original = sOTHERINFO
    DoConsiderForJoining Text:=sOTHERINFO, Field:=Me!REMARKS2, PrefixPattern:="Witnesses:"
    If sOTHERINFO <> sOTHERINFO_Original Then
        Let Me!OTHERINFO.Value = "Witnesses: " & sOTHERINFO
    End If
    
    Let sAUTHORITY = Nz(Me!AUTH.Value)
    Let sAUTHORITY_Original = sAUTHORITY
    DoConsiderForJoining Text:=sAUTHORITY, Field:=Me!REMARKS2, PrefixPattern:="Auth(ority)?\s*con[']?t([']?d)?[.]?[:]?"
    If sAUTHORITY <> sAUTHORITY_Original Then
        Let Me!AUTH.Value = sAUTHORITY
        HighlightField Field:=Me!AUTH
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
    
    Controls.Add Item:=Me!LNAME
    Controls.Add Item:=Me.FNAME
    Controls.Add Item:=Me.MI
        
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
    
    Let Me!slug.Value = Join(Delimiter:="-", List:=Elements)
    If Nz(Me!slug.Value) <> Nz(Me!slug.OldValue) Then
        HighlightField Field:=Me!slug
    End If
    
    Set Elements = Nothing
    Set Controls = Nothing
End Sub

Private Sub DoResetHighlightedFields()
    Dim Field As Variant
    For Each Field In Me.Controls
        If TypeName(Field) = "TextBox" Then
            HighlightField Field:=Field, Tag:=""
        End If
    Next Field
End Sub
