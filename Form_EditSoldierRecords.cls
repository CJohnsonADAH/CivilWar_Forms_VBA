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

Private iPreviousRecord As Long
Private sPreviousName As String
Private sPreviousUnit As String
Private bPreviouslyFlagged As Boolean
Private bHighlightNames As Boolean
    
Private Property Get HighlightNames() As Boolean
    Let HighlightNames = bHighlightNames
End Property

Private Property Let HighlightNames(RHS As Boolean)
    Let bHighlightNames = RHS
End Property

Private Property Get CRLF() As String
    Let CRLF = Chr$(13) & Chr$(10)
End Property

Private Property Get PreviousRecord() As Long
    Let PreviousRecord = iPreviousRecord
End Property
Private Property Let PreviousRecord(RHS As Long)
    Let iPreviousRecord = RHS
End Property
Private Property Get PreviousName() As String
    Let PreviousName = sPreviousName
End Property
Private Property Let PreviousName(RHS As String)
    Let sPreviousName = RHS
End Property
Private Property Get PreviousUnit() As String
    Let PreviousUnit = sPreviousUnit
End Property
Private Property Let PreviousUnit(RHS As String)
    Let sPreviousUnit = RHS
End Property
Private Property Get PreviouslyFlagged() As Boolean
    Let PreviouslyFlagged = bPreviouslyFlagged
End Property
Private Property Let PreviouslyFlagged(RHS As Boolean)
    Let bPreviouslyFlagged = RHS
End Property

Private Function EngagementsBreaker() As String
    Let EngagementsBreaker = "^(Killed|Ab|C(apt|[.])|Pr(es|[.:\s])|Sev)"
End Function

Private Sub AUTH_DblClick(Cancel As Integer)
    Dim sAuth As String
    Dim sAuthRegex As String
        
    Let sAuth = Me!AUTH.Value
    Let sAuthRegex = AuthorityContdRegex
    
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
        Let sEngage2 = RegexReplace(sEngage2, "^[(]?Remarks( con[']?t)?[:)]?\s*", "")
            
        Let sEngage = FieldsJoined(sEngage, sEngage2)
        Let Me!ENGAGE2.Value = Null
        HighlightField Field:=Me!ENGAGE2
    End If
        
    Let Me!REMARKS.Value = RegexReplace(Me!REMARKS.Value, "^[(]?(Con[']?t)[)]?[:]", "")
    
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

Private Sub cmdFilterNullSlug_Click()
    If Me.FilterOn Then
        Let Me.FilterOn = False
    Else
        Let Me.Filter = "Len(Nz([slug])) = 0 AND Len(Nz(REGUNIT)) > 0"
        Let Me.FilterOn = True
    End If
    
End Sub

Private Sub cmdJoinEm_Click()
    DoProcessENGAGE_Field
    DoProcessREMARKS_Field
End Sub

Private Sub cmdLinkPossiblyDuplicates_Click()
    Dim aAliases() As String
    Dim dAliases As New Dictionary
    Dim sAliasRecNumber As String
    Dim I As Long
    Dim I0 As Long
    Dim N As Long
    Dim sId As String
    
    Let sAliasRecNumber = Nz(Me!alias_recnumber.Value)
    Let aAliases = RegexSplit(Text:=sAliasRecNumber, Pattern:="\s*;\s*")
    
    Let I0 = LBound(aAliases)
    Let N = UBound(aAliases)
    For I = I0 To N
        Let sId = Trim(aAliases(I))
        If Len(sId) > 0 Then
            dAliases.Item(sId) = sId
        End If
    Next I
    
    dAliases.Item(Trim(Me!recnumber.Value)) = Trim(Me!recnumber.Value)
    dAliases.Item(Trim(Me!txtPossiblyDuplicatesId2.Value)) = Trim(Me!txtPossiblyDuplicatesId2.Value)
    
    Dim SortedKeys() As Variant
    
    Let SortedKeys = dAliases.Keys
    BubbleSortList SortedKeys
    
    Let Me!alias_recnumber.Value = Join("; ", SortedKeys)
    
    Set dAliases = Nothing
    
End Sub

Private Sub cmdNextRecord_Click()
    DoCmd.GoToRecord Record:=acNext
End Sub

Private Sub cmdPrevRecord_Click()
    DoCmd.GoToRecord Record:=acPrevious
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
        
        If cUnitCo.Count <= 1 Then
            Let sUnitCo = GetAuthoritySlug(Text:=Nz(Me!AUTH.Value))
            If Len(sUnitCo) > 0 Then
                cUnitCo.Add sUnitCo
            End If
        End If
        
        If cUnitCo.Count <= 1 Then
            cUnitCo.Add VBA.Format$(Me!recnumber.Value, "0")
        End If
        
        Let sUnitCo = Join("-", List:=cUnitCo)
        Let Me!slug.Value = sUnitCo
        
        'Set cUnitCo = Null
        
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
    DoResetUtilityFields
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Dim oHistory As cRevisedTable
    Dim Rs As DAO.Recordset
        
    If Me.Dirty Then
        Set oHistory = New cRevisedTable
        Let oHistory.IDName = "recnumber"

        Set Rs = Me.RecordsetClone
        On Error GoTo PauseForASec
        oHistory.DoStoreRevision ID:=Me!recnumber.Value, Rs:=Rs
        On Error GoTo 0
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
    
PauseForASec:
    If Err.Number = 3027 Then
        Debug.Print "Got an ERROR: ", Err.Number, Err.Description, "... try again in a sec."
    
        Dim I As Integer
        Dim N As Integer
        Let N = 255
        For I = 0 To (N - 1)
            Let Me!recnumber.BackColor = RGB(255, 255 * (I / N), 255 * (I / N))
            DoEvents
        Next I
        Let Me!recnumber.BackColor = RGB(255, 255, 255)
    
        oHistory.DoStoreRevision ID:=Me!recnumber.Value, Rs:=Rs
        Resume Next
    End If
    
End Sub

Private Sub DoCheckForPossibleDuplicates()
    Dim sCurName As String
    Dim sPrevSlug As String
    Dim sCurSlug As String
    Dim sCurUnit As String
        
    Let sCurName = Me!FNAME.Value & " " & Me!MI.Value & " " & Me!LNAME.Value
    Let sCurUnit = Me!BRANCH.Value & " " & Me!REGUNIT.Value & " Co. " & Me!COMUNIT.Value & " (" & Me!COMUNITNA.Value & ")"
    
    Let sCurSlug = RegexReplace(Pattern:="(^[^a-z0-9]+|[^a-z0-9]+$)", Value:=LCase(Trim(sCurName)), Replace:="")
    Let sCurSlug = RegexReplace(Pattern:="[^a-z0-9]+", Value:=sCurSlug, Replace:="-")
    Let sPrevSlug = RegexReplace(Pattern:="(^[^a-z0-9]+|[^a-z0-9]+$)", Value:=LCase(Trim(PreviousName)), Replace:="")
    Let sPrevSlug = RegexReplace(Pattern:="[^a-z0-9]+", Value:=sPrevSlug, Replace:="-")

    Dim TestHighlight As Boolean
    
    Let TestHighlight = False
    Debug.Print "Previous Name:", PreviousName & " (" & sPrevSlug & ")", "Current Name:", sCurName & " (" & sCurSlug & ")"
    
    Let Me!txtPossiblyDuplicatesName.Height = 0
    Let Me!txtPossiblyDuplicatesName2.Height = 0
    Let Me!txtPossiblyDuplicatesName.Value = Null
    Let Me!txtPossiblyDuplicatesName2.Value = Null
    
    If TestHighlight Or ((sPrevSlug = sCurSlug) And (PreviousRecord <> Me!recnumber.Value)) Then
        Let HighlightNames = True
        
        Let Me!txtPossiblyDuplicatesName.Height = 400
        Let Me!txtPossiblyDuplicatesName2.Height = 400
        Let Me!txtPossiblyDuplicatesName.Value = sCurName
        Let Me!txtPossiblyDuplicatesUnitCo.Value = sCurUnit
        Let Me!txtPossiblyDuplicatesName2.Value = PreviousName
        Let Me!txtPossiblyDuplicatesUnitCo2.Value = PreviousUnit
        Let Me!txtPossiblyDuplicatesId2.Value = PreviousRecord

        Let PreviouslyFlagged = True
    
    Else
        Let PreviouslyFlagged = False
    End If
    
    Let PreviousName = sCurName
    Let PreviousUnit = sCurUnit
    Let PreviousRecord = Nz(Me!recnumber.Value)

End Sub

Private Sub Form_Current()
    Let HighlightNames = False
    
    Let Me!txtRemarksCandidates.Value = Null
    Let Me!txtRemarksCandidatesNo.Value = Null

    DoCheckForPossibleDuplicates
    
    Let Me!RevisedTimestamp.Locked = True
    Let Me!RevisedBy.Locked = True
    
    Let Me!REMARKS2.Visible = (Len(Nz(Me!REMARKS2.Value)) > 0)
    Let Me!REMARKS2.Height = IIf(Me!REMARKS2.Visible, 810, 10)

    DoResetHighlightedFields
    If Len(Nz(Me!slug.Value)) = 0 Then
        DoFillInSlug
    End If
    
    If TypeName(Me!REMARKS2.Value) = "String" Then
        If ShouldProcessREMARKS2 Then
            DoProcessREMARKS2
        End If
    End If
    
    If HighlightNames Then
        HighlightField Field:=Me!FNAME
        HighlightField Field:=Me!LNAME
        HighlightField Field:=Me!MI
        HighlightField Field:=Me!alias_recnumber
    
        HighlightField Field:=Me!txtPossiblyDuplicatesName
        HighlightField Field:=Me!txtPossiblyDuplicatesName2
        HighlightField Field:=Me!txtPossiblyDuplicatesUnitCo
        HighlightField Field:=Me!txtPossiblyDuplicatesUnitCo2
        HighlightField Field:=Me!txtPossiblyDuplicatesId2
    
    End If
End Sub


Private Sub Form_Undo(Cancel As Integer)
    DoResetUtilityFields
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

Private Sub LNAME_DblClick(Cancel As Integer)
    Dim RecordNumber As String
    
    If HasRecordNumber(Nz(Me!LNAME.Value)) Then
        Let RecordNumber = GetRecordNumber(Field:=Nz(Me!LNAME.Value))
        Let LNAME.Value = StripRecordNumber(Field:=Nz(Me!LNAME.Value))
        
        If Len(RecordNumber) > 0 Then
            Let RecordNumber = "(Record " & RecordNumber & ".)"
            Let Me!OTHERINFO.Value = RecordNumber & Chr$(13) & Chr$(10) & Nz(Me!OTHERINFO.Value)
            
            HighlightField Field:=Me!LNAME
            HighlightField Field:=Me!OTHERINFO
        End If
        
    End If
    
End Sub

Private Sub REMARKS_DblClick(Cancel As Integer)
    DoProcessREMARKS_Field
End Sub

Private Sub DoCycleREMARKSCandidates()
    Dim aCandidates() As String
    Dim I As Long
    Dim Range As Long
    
    If Len(Nz(Me!txtRemarksCandidates.Value)) > 0 Then
        Let aCandidates = RegexSplit(Text:=Nz(Me!txtRemarksCandidates.Value), Pattern:=CRLF & "--- " & CRLF)
        Let I = CInt(Nz(Me!txtRemarksCandidatesNo.Value))
        Let Range = (UBound(aCandidates) - LBound(aCandidates)) + 1
    
        Let I = (I Mod Range)
    
        Let Me!REMARKS.Value = aCandidates(I)
        HighlightField Field:=Me!REMARKS
    
        Let Me!txtRemarksCandidatesNo.Value = VBA.Format$(I + 1, "0")
    End If
    
End Sub

Private Sub DoProcessREMARKS_Field()
    Dim sEngage2 As String
    Dim sRemarksRegex As String
    
    Dim sRemarks As String
    Let sRemarks = Nz(Me!REMARKS.Value)
    
    If Len(Nz(Me!txtRemarksCandidates.Value)) > 0 Then
        DoCycleREMARKSCandidates
        Exit Sub
    End If
    
    DoConsiderForJoining Text:=sRemarks, Field:=Me!REMARKS2, PrefixPattern:="(R+emarks\s*" & ContdRegex & ")?", Alternatives:=Me!txtRemarksCandidates
    Let Me!txtRemarksCandidatesNo.Value = "1"
    
    If sRemarks <> Nz(Me!REMARKS.Value) Then
        Let Me!REMARKS.Value = sRemarks
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

Public Function ShouldProcessREMARKS2() As Boolean
    Let ShouldProcessREMARKS2 = ( _
        RegexMatch(Value:=Me!REMARKS2.Value, Pattern:=AuthorityAutoProcessRegex) _
        Or (Nz(Me!AUTH.Value) Like "*Remarks con't*" _
        Or Nz(Me!AUTH.Value) Like "*See Remarks cont'd*" _
        Or Nz(Me!REMARKS2.Value) Like "* Authority sources* have * information.") _
    )
End Function

Public Sub DoProcessREMARKS2()
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
    
    Let sOTHERINFO = Nz(Me!OTHERINFO.Value)
    Let sOTHERINFO_Original = sOTHERINFO
    DoConsiderForJoining Text:=sOTHERINFO, Field:=Me!REMARKS2, PrefixPattern:="Pension\b"
    If sOTHERINFO <> sOTHERINFO_Original Then
        Let Me!OTHERINFO.Value = "Pension" & sOTHERINFO
    End If
    
    Let sOTHERINFO = Nz(Me!OTHERINFO.Value)
    Let sOTHERINFO_Original = sOTHERINFO
    DoConsiderForJoining Text:=sOTHERINFO, Field:=Me!REMARKS2, PrefixPattern:="Address\b"
    If sOTHERINFO <> sOTHERINFO_Original Then
        Let Me!OTHERINFO.Value = "Address" & sOTHERINFO
    End If
    
    Let sOTHERINFO = Nz(Me!OTHERINFO.Value)
    Let sOTHERINFO_Original = sOTHERINFO
    DoConsiderForJoining Text:=sOTHERINFO, Field:=Me!REMARKS2, PrefixPattern:="Residence\b"
    If sOTHERINFO <> sOTHERINFO_Original Then
        Let Me!OTHERINFO.Value = "Residence" & sOTHERINFO
    End If
    
    Let sAUTHORITY = Nz(Me!AUTH.Value)
    Let sAUTHORITY_Original = sAUTHORITY
    If sAUTHORITY Like "*Remarks con't*" Or sAUTHORITY Like "*See Remarks cont'd*" Then
        Let sAUTHORITY = Replace(Expression:=sAUTHORITY, Find:="See Remarks con't to allow for space needed:", Replace:="")
        Let sAUTHORITY = Replace(Expression:=sAUTHORITY, Find:="(See Remarks cont'd)", Replace:="")
    End If
    
    Dim sSeeRemarksRegex As String
    Let sSeeRemarksRegex = "(" & _
        "\s*[^A-Z0-9a-z]+(see\s*)?remark[s]?\s*con[']?[dt][^A-Z0-9a-z]*\s*" & _
        "|" & _
        "[^A-Za-z0-9]con[']?t[.]? in Re?ma?r?ks?[^A-Za-z0-9]" & _
        ")$"
    If RegexMatch(Value:=sAUTHORITY, Pattern:=sSeeRemarksRegex) Then
        Let sAUTHORITY = RegexReplace(Value:=sAUTHORITY, Pattern:=sSeeRemarksRegex, Replace:="")
    End If
    
    Dim vRemarksParts
    If RegexMatch(Value:=Nz(Me!REMARKS2.Value), Pattern:=AuthoritySplitRegex) Then
        Let vRemarksParts = RegexSplit(Text:=Nz(Me!REMARKS2.Value), Pattern:=AuthoritySplitRegex)
        
        Dim S0 As String, SN As String
        Dim i_0 As Long, i_N As Long
        Let i_0 = LBound(vRemarksParts)
        Let i_N = UBound(vRemarksParts)
        If i_N >= i_0 Then
            Let SN = vRemarksParts(i_N)
            Let S0 = vRemarksParts(i_0)
            
            Let sAUTHORITY = FieldsJoined(Text1:=sAUTHORITY, Text2:=SN)
            
            Let Me!REMARKS2.Value = S0
            Let Me!REMARKS.Value = FieldsJoined(Text1:=Nz(Me!REMARKS.Value), Text2:=Nz(Me!REMARKS2.Value))
            Let Me!REMARKS2.Value = ""
            
            HighlightField Field:=Me!REMARKS
            HighlightField Field:=Me!REMARKS2
            HighlightField Field:=Me!AUTH
            
        End If
    
    End If
    
    DoConsiderForJoining Text:=sAUTHORITY, Field:=Me!REMARKS2, PrefixPattern:=AuthorityPrefixRegex
    If sAUTHORITY <> sAUTHORITY_Original Then
        Let Me!AUTH.Value = sAUTHORITY
        HighlightField Field:=Me!AUTH
    End If
End Sub

Private Sub REMARKS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iLeftButton As Long
    Dim iMidButton As Long
    Dim bitShift As Long
    
    Let iLeftButton = 1
    Let iMidButton = 4
    Let bitShift = 1
    
    Debug.Print "Button=", Button, "Shift=", Shift
    If Button = iLeftButton And (Shift And bitShift) Then
        DoCycleREMARKSCandidates
    ElseIf Button = iMidButton Then
        DoCmd.GoToRecord Record:=acNext
    End If
End Sub

Private Sub REMARKS2_DblClick(Cancel As Integer)
    DoProcessREMARKS2
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
    
    If RegexMatch(Me!LNAME.Value, "^Slave") Then
        Controls.Add Item:=Me!FNAME
        Controls.Add Item:=Me!MI
        Controls.Add Item:=Me!LNAME
    Else
        Controls.Add Item:=Me!LNAME
        Controls.Add Item:=Me!FNAME
        Controls.Add Item:=Me!MI
    End If
    
    For Each Control In Controls
        Dim I As Long
        Dim ControlWords() As String
        Dim ControlText As String
        
        Let ControlText = FilterName(Nz(Control.Value))
        Let ControlWords = RegexSplit(Text:=ControlText, Pattern:="[^A-Z0-9]+")
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
    
    If Len(Nz(Me!REGUNIT.Value)) > 0 Then
        DoElaborateAlphanumSlug
    End If
    
End Sub

Private Sub DoResetUtilityFields()
    Let Me!txtRemarksCandidates.Value = Null
    Let Me!txtRemarksCandidatesNo.Value = Null
End Sub

Private Sub DoResetHighlightedFields()
    Dim Field As Variant
    For Each Field In Me.Controls
        If TypeName(Field) = "TextBox" Then
            HighlightField Field:=Field, Tag:=""
        End If
    Next Field
End Sub
