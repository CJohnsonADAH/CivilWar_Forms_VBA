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
    
    Let sText1 = RegexReplace(Value:=sText1, Pattern:="\s*[(]\s*(con[']?t[']?d|con[']?t|continued?)\s*[)]\s*$", Replace:="")
    
    Let sText1 = RegexReplace(Value:=sText1, Pattern:="^\s*((Lost)\s*(use\s*of|an|a\b)?\s*(left|right)?\s*(arm|leg|hand|hip|hearing))\s*[.]?\s*$", Replace:="""$1.""")
    Let Field1Filter = sText1
End Function

Function Field2Filter(ByVal Text2 As String)
    Dim sText2 As String
    
    Let sText2 = Text2
    Let sText2 = RegexReplace(Value:=sText2, Pattern:="^\s*[(]cont([']d)?[)]\s*", Replace:="")
    
    Let Field2Filter = sText2
End Function

Function FieldsJoined(ByVal Text1 As String, ByVal Text2 As String, Optional ByVal Breaker As String, Optional ByRef Alternatives As Variant) As String
    Dim sFieldsJoined As String
    Dim sAlternatives As String
    Dim cAlternatives As New Collection
    Dim sChar As String
    Dim sRest As String
    
    Dim CRLF As String
    
    Let CRLF = Chr$(13) & Chr$(10)

    Let sFieldsJoined = Field1Filter(Text1) & FieldJoiner(Text1, Text2, Breaker) & Field2Filter(Text2)
    If Not IsMissing(Alternatives) Then
        cAlternatives.Add sFieldsJoined
        
        If RegexMatch(Text2, "^\s*[A-Z]", MatchCase:=True) Then
            Let sRest = LTrim(Text2)
            Let sChar = Left(sRest, 1)
            Let sRest = Right(sRest, Len(sRest) - 1)
            
            Let sFieldsJoined = RTrim(Field1Filter(Text1)) & FieldJoiner(Text1, Text2, Breaker) & LTrim(Field2Filter(LCase(sChar) & sRest))
            cAlternatives.Add sFieldsJoined
        End If
        
        Let sFieldsJoined = RTrim(Field1Filter(Text1)) & ". " & LTrim(Field2Filter(Text2))
        cAlternatives.Add sFieldsJoined
        
        Let sFieldsJoined = RTrim(Field1Filter(Text1)) & "; " & LTrim(Field2Filter(Text2))
        cAlternatives.Add sFieldsJoined
        
        Let sFieldsJoined = RTrim(Field1Filter(Text1)) & " [...] " & LTrim(Field2Filter(Text2))
        cAlternatives.Add sFieldsJoined
        
        Let sFieldsJoined = cAlternatives.Item(1)
    End If
    
    If cAlternatives.Count > 0 Then
        If Not IsMissing(Alternatives) Then
            Let Alternatives.Value = Join(CRLF & "--- " & CRLF, cAlternatives)
        End If
    End If
    Set cAlternatives = Nothing
    Let FieldsJoined = sFieldsJoined
End Function

Sub DoConsiderForJoining(ByRef Text As String, ByRef Field As Variant, ByVal PrefixPattern As String, Optional ByVal Breaker As String, Optional ByRef Alternatives As Variant)
    Dim vField As Variant
    Dim sRegex As String
    
    If Len(Nz(Field.Value)) > 0 Then
        Let vField = Field.Value
        Let sRegex = "^([(>]|\[)?(" & PrefixPattern & ")([)]|\])?"
        If RegexMatch(Value:=Nz(vField), Pattern:=sRegex) Then
            Let vField = RegexReplace(Nz(vField), sRegex, "")
            Let Text = FieldsJoined(Text, vField, Breaker:=Breaker, Alternatives:=Alternatives)
            Let Field.Value = ""
            HighlightField Field:=Field
        End If
    End If

End Sub

Sub DoJoinEngagementFields(ByRef RunningText As String, ByRef ENGAGE As Variant, ByRef ENGAGE2 As Variant)
    Dim sEngage2 As String
    
    Let RunningText = RegexReplace(Value:=RunningText, Pattern:="[(]Con[']?t([:.;,]|[']?d)?[)]$", Replace:="")
    Let RunningText = RegexReplace(Value:=RunningText, Pattern:="[(]\s*continued\s*[)]$", Replace:="")

    Let sEngage2 = Nz(ENGAGE2.Value)
    If RegexMatch(RunningText, Pattern:="^\s*[(]Remarks(\s*begin)?[)]") Then
        Let sEngage2 = RegexReplace(Value:=sEngage2, Pattern:="^\s*[(]Remarks(\s*begin)?[)]\s*", Replace:="")
    End If
    Let sEngage2 = RegexReplace(Value:=sEngage2, Pattern:="^[(]Con[']?t([:.;,]|[']?d)?[)]", Replace:="")
    Let sEngage2 = RegexReplace(Value:=sEngage2, Pattern:="^[(]\s*continued\s*[)]", Replace:="")
    
    Let sEngage2 = RegexReplace(Value:=sEngage2, Pattern:=EngagementsContdCombinedRegex(Prefix:="^\s*[(]?\s*", Suffix:="\s*([)][:.]?|[:.])\s*"), Replace:="")
    Debug.Print "REGEX:", EngagementsContdCombinedRegex(Prefix:="^")
    
    Let sEngage2 = RegexReplace(Value:=sEngage2, Pattern:="^\s*[(][^)]+con[']?t(inu|['])?e?[']?d(\s*[^)]+)?[)]\s*", Replace:=" ")
    Let sEngage2 = RegexReplace(Value:=sEngage2, Pattern:="^\s*[(]\s*cont[.]\s*[^)]*[)]\s*", Replace:=" ")
    Let sEngage2 = RegexReplace(Value:=sEngage2, Pattern:="^\s*[(]\s*continued\s*[)]\s*", Replace:="")

    Let sEngage2 = RegexReplace(Value:=sEngage2, Pattern:="\s*[(]\s*con[']?t([']?d)?\s*[)]\s*$", Replace:="")
    Let sEngage2 = RegexReplace(Value:=sEngage2, Pattern:="\s*[(]\s*(con[']?t([']?d)?|see)\s*remarks\s*con[']?t([']?d)?\s*[)]\s*$", Replace:="")
    
    Let RunningText = FieldsJoined(RunningText, sEngage2)
    Let RunningText = RegexReplace(Value:=RunningText, Pattern:="\s+", Replace:=" ")

    Let RunningText = RegexReplace(Value:=RunningText, Pattern:="(;)(\S)", Replace:="$1 $2")
    Let RunningText = RegexReplace(Value:=RunningText, Pattern:="([A-Za-z])([0-9]{2,4})", Replace:="$1 $2")
    
    Let ENGAGE2.Value = Null
    HighlightField Field:=ENGAGE2
End Sub

Sub DoFilterREMARKSField(ByRef Field As Variant)
    Dim CRLF As String
    
    Let CRLF = Chr$(13) & Chr$(10)
    
    If Not IsNull(Field.Value) Then
        Let Field.Value = RegexReplace(Value:=Field.Value, Pattern:="^(P[.]?O[.]?|Post\s+Office|Residence|Res[.]?)(\s+|\s*[-;,]\s*)([A-Z])", Replace:="$1: $3")
        Let Field.Value = RegexReplace(Value:=Field.Value, Pattern:="^\s*[(]Remarks(\s*begin)?[)]\s*", Replace:="")
        Let Field.Value = RegexReplace(Value:=Field.Value, Pattern:="\s{7,}", Replace:=CRLF & CRLF)
    End If
End Sub

Function ContdRegex() As String
    Let ContdRegex = "\b(con[.']?t?(in(ue)?)?[.']?d?[.']?|con['.]?t['.]?in['.]?((ue)?d))[.]?"
End Function

Function EngagementsContdRegexes() As Collection
    Dim sReSep As String
    Dim sReContd As String
    
    Let sReSep = "\s*([.,;:!\-]*\s*)*"
    Let sReContd = ContdRegex
    
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

Function FilterName(ByVal Text As String)
    Let FilterName = Trim(StripRecordNumber(Field:=Text))
End Function

Function FieldRecordNumberRegex() As String
    Let FieldRecordNumberRegex = "\s*[(]?(\b([0-9]+)\s*((of|[/])\s*([0-9]+)\b)?)\s*[)]?\s*$"
End Function

Function StripRecordNumber(ByVal Field As String) As String
    Let StripRecordNumber = RegexReplace(Value:=Field, Pattern:=FieldRecordNumberRegex, Replace:="")
End Function

Function GetRecordNumber(ByVal Field As String) As String
    Dim dBits As Dictionary
    Dim sOrdinal As String
    Dim sTotal As String
    
    Set dBits = RegexComponents(Value:=Field, Pattern:=FieldRecordNumberRegex)
    Let sOrdinal = dBits.Item(2)
    Let sTotal = dBits.Item(5)
    
    If Len(sTotal) > 0 Then
        Let GetRecordNumber = sOrdinal & " of " & sTotal
    Else
        Let GetRecordNumber = sOrdinal
    End If
    
End Function

Function HasRecordNumber(ByVal Field As String) As String
    Let HasRecordNumber = RegexMatch(Value:=Field, Pattern:=FieldRecordNumberRegex)
End Function

Function GetAuthoritySlug(ByVal Text As String) As String

'Pattern 1
'Letter, The Adjutant General, 1914/07/11 to H. Y. Brooke

    Dim sDate As String
    Dim sreUSAGO As String
    Dim sreUSAGOMid As String
    Dim sreDate As String
    
    Let sreUSAGO = "(Adj([.]|utant)?\s*General([']?s)?(\s*(Office|Ofc))?|USAGO)[^a-z0-9]*\s*"
    Let sreUSAGOMid = "(Washington[,]?\s*D[.]?\s*C[.]?|Department\s*of\s*the\s*Army|letter)?[^a-z0-9]?\s*"
    Let sreDate = "([0-9]{1,4}/[0-9]{1,}/[0-9]{1,4})"
    
    Let sDate = RegexComponent(Value:=Text, Pattern:=sreUSAGO & sreUSAGOMid & sreDate, Part:=7)
    If Len(sDate) > 0 Then
        Let sDate = RegexReplace(Value:=sDate, Pattern:="[^0-9]+", Replace:="")
        Let GetAuthoritySlug = "usago-" & sDate
    End If
    
End Function

Function AuthorityWordRegex() As String
    Let AuthorityWordRegex = "Auth(or?(i?ty)?)?"
End Function

Function AuthoritySplitRegex() As String
    Let AuthoritySplitRegex = "(" & _
        "([^A-Za-z0-9]|^)\s*" & ContdRegex & "\s+from\s+" & AuthorityWordRegex & "([^A-Za-z0-9]+)+\s*" & _
        "|" & _
        "([^A-Za-z0-9]|^)\s*" & AuthorityWordRegex & "[^A-Z0-9]*(con[']?t(inue)?[']?(d)?)[^A-Z0-9]*([^A-Za-z0-9]+)+\s*" & _
        ")"
        
End Function

Function AuthorityContdRegex() As String
    Let AuthorityContdRegex = AuthorityWordRegex & "\s*cont[.]|Auth(or?(ity)?)?(\s+" & ContdRegex & "?[:.]?\s*"
End Function

Function AuthorityPrefixRegex() As String
    Dim sReContd As String
    Let sReContd = ContdRegex
    Let AuthorityPrefixRegex = "([^A-Z0-9]*Author?ity[)]+|" & sReContd & "[.]?\s*[^A-Z0-9]?Auth(or?ity)?[^A-Z0-9]?\s*|[""]?Auth(or?ity)?\s*[^A-Z0-9]*\s*" & sReContd & "\s*[^A-Z0-9]*\s*)?[.]?[:]?"
End Function

Function AuthorityAutoProcessRegex() As String
    Let AuthorityAutoProcessRegex = "^[^A-Za-z0-9]*(Con\S*\s*[^A-Za-z0-9]*(from\s*)?[Aa]uth(ority)?|[Aa]uth(ority)?\s*[^A-Za-z0-9]*\s*Con\S*)"
End Function
