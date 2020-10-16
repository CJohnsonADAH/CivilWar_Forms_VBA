Attribute VB_Name = "modCw"
Option Compare Database
Option Explicit

Public Sub SaveMe3()
    Dim Qdf As DAO.QueryDef
    Dim Rs As DAO.Recordset
    Dim o As New cRevisedTable
    
    Let o.IDName = "recnumber"
    
    Set Qdf = CurrentDb.CreateQueryDef(Name:="")
    Let Qdf.SQL = "SELECT * FROM Cw WHERE recnumber=[paramRecnumber]"
    Let Qdf.Parameters("paramRecnumber") = 106234
    
    Set Rs = Qdf.OpenRecordset
    
    o.DoStoreRevision ID:=Rs!recnumber.Value, Rs:=Rs
    
    Rs.Close: Set Rs = Nothing
    Qdf.Close: Set Qdf = Nothing
    

End Sub

Sub HighlightField(ByRef Field As Variant, Optional ByVal Tag As Variant)
    If IsMissing(Tag) Then
        Let Tag = "Modified"
    End If
    
    If Len(Tag) > 0 Then
        Let Field.BackColor = RGB(255, 255, 192)
    Else
        Let Field.BackColor = RGB(255, 255, 255)
    End If
    
    Let Field.Tag = Tag
End Sub

Function FieldJoiner(ByVal Text1 As String, ByVal Text2 As String, Optional ByVal Breaker As String) As String
    Dim sSep As String
    
    Let sSep = ""
    If Len(Trim(Text1)) > 0 Then
        Let sSep = " "
        If (Len(Breaker) > 0) And RegexMatch(Value:=Text2, Pattern:=Breaker) Then
            If Not RegexMatch(Value:=Text1, Pattern:="[;.,]\s*$") Then
                Let sSep = ". "
            End If
        End If
    End If

    Let FieldJoiner = sSep
End Function

Function Field1Filter(ByVal Text1 As String)
    Dim sText1 As String
    
    Let sText1 = Text1
    Let sText1 = RegexReplace(Value:=sText1, Pattern:="\s*([.]\s*){2,}\s*$", Replace:="")
    
    Let Field1Filter = sText1
End Function

Function Field2Filter(ByVal Text2 As String)
    Dim sText2 As String
    
    Let sText2 = Text2
    Let sText2 = RegexReplace(Value:=sText2, Pattern:="^\s*[(]cont([']d)?[)]\s*", Replace:="")
    
    Let Field2Filter = sText2
End Function

Function FieldsJoined(ByVal Text1 As String, ByVal Text2 As String, Optional ByVal Breaker As String) As String
    Let FieldsJoined = Field1Filter(Text1) & FieldJoiner(Text1, Text2, Breaker) & Field2Filter(Text2)
End Function

Sub DoConsiderForJoining(ByRef Text As String, ByRef Field As Variant, ByVal PrefixPattern As String, Optional ByVal Breaker As String)
    Dim vRemarks As Variant
    Dim sRegex As String
    
    If Len(Nz(Field.Value)) > 0 Then
        Let vRemarks = Field.Value
        Let sRegex = "^([(>]|\[)?(" & PrefixPattern & ")([)]|\])?"
        If RegexMatch(Value:=Nz(vRemarks), Pattern:=sRegex) Then
            Let vRemarks = RegexReplace(Nz(vRemarks), sRegex, "")
            Let Text = FieldsJoined(Text, vRemarks, Breaker:=Breaker)
            Let Field.Value = ""
            HighlightField Field:=Field
        End If
    End If

End Sub

Sub DoJoinEngagementFields(ByRef RunningText As String, ByRef ENGAGE As Variant, ByRef ENGAGE2 As Variant)
    Dim sEngage2 As String
    
    Let RunningText = RegexReplace(Value:=RunningText, Pattern:="[(]Cont([:.;,]|[']?d)?[)]$", Replace:="")
    Let RunningText = RegexReplace(Value:=RunningText, Pattern:="[(]\s*continued\s*[)]$", Replace:="")

    Let sEngage2 = Nz(ENGAGE2.Value)
    If RegexMatch(RunningText, Pattern:="^\s*[(]Remarks(\s*begin)?[)]") Then
        Let sEngage2 = RegexReplace(Value:=sEngage2, Pattern:="^\s*[(]Remarks(\s*begin)?[)]\s*", Replace:="")
    End If
    
    Let sEngage2 = RegexReplace(Value:=sEngage2, Pattern:=EngagementsContdCombinedRegex(Prefix:="^\s*[(]?\s*", Suffix:="\s*([)][:.]?|[:.])\s*"), Replace:="")
    Debug.Print "REGEX:", EngagementsContdCombinedRegex(Prefix:="^")
    
    Let sEngage2 = RegexReplace(Value:=sEngage2, Pattern:="^\s*[(][^)]+con[']?t(inu|['])?e?[']?d(\s*[^)]+)?[)]\s*", Replace:=" ")
    Let sEngage2 = RegexReplace(Value:=sEngage2, Pattern:="^\s*[(]\s*cont[.]\s*[^)]*[)]\s*", Replace:=" ")
    Let sEngage2 = RegexReplace(Value:=sEngage2, Pattern:="^\s*[(]\s*continued\s*[)]\s*", Replace:="")

    Let RunningText = FieldsJoined(RunningText, sEngage2)
    Let RunningText = RegexReplace(Value:=RunningText, Pattern:="\s+", Replace:=" ")

    
    Let ENGAGE2.Value = Null
    HighlightField Field:=ENGAGE2
End Sub

Sub DoFilterREMARKSField(ByRef Field As Variant)
    If Not IsNull(Field.Value) Then
        Let Field.Value = RegexReplace(Value:=Field.Value, Pattern:="^(P[.]?O[.]?|Post\s+Office|Residence|Res[.]?)(\s+|\s*[-;,]\s*)([A-Z])", Replace:="$1: $3")
        Let Field.Value = RegexReplace(Value:=Field.Value, Pattern:="^\s*[(]Remarks(\s*begin)?[)]\s*", Replace:="")
    End If
End Sub

Function EngagementsContdRegexes() As Collection
    Dim sReSep As String
    Dim sReContd As String
    
    Let sReSep = "\s*([.,;:!\-]*\s*)*"
    Let sReContd = "\b(con[.']?d|con[.']?t|con[.']?t[.']?d|contin[.']?d|continued)[.]?"
    
    Set EngagementsContdRegexes = Coll( _
            "\bEngag(e(ment(s)?)?)?\b" & sReSep & sReContd, _
            sReContd & sReSep & "\bEngag(e(ment(s)?)?)?\b" _
    )
End Function

Function EngagementsContdCombinedRegex(Optional ByVal Prefix As String, Optional ByVal Suffix As String) As String
   
    If Len(Prefix) = 0 Then
        Let Prefix = ""
    End If
    
    If Len(Suffix) = 0 Then
        Let Suffix = "[:.]?\s*"
    End If
    
    Let EngagementsContdCombinedRegex = Prefix & "(" & Join("|", List:=EngagementsContdRegexes) & ")" & Suffix
End Function

