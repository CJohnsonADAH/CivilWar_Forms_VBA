Attribute VB_Name = "modMyVBA"
'**
'* modMyVBA: Fluent routine problem-solving and more bearable development processes in VBA.
'*
'* @author C. Johnson
'* @version 2018.1204
'*
'* @uses Scripting.Dictionary {420B2830-E718-11CF-893D-00A0C9054228}/1/0
'*
'* @doc README.txt MyVBA
'* MyVBA is a collection of tools I use across a whole bunch of projects to make routine problem-solving in VBA
'* just a bit quicker, more natural and make the code more readable. Some of the tools herein are mainly to help
'* with debugging, version control, and other development tasks within the VBA IDE. Other tools are here to make
'* VBA coding somewhat more like coding in other dynamic scripting languages.
'*
'* == modMyVBA ==
'* modMyVBA provides toolbox of functions for easing work with common data structures like lists and strings,
'* as well as some frequently used patterns for storing persistent settings in an Access project, etc.
'*
'* == modDevTools ==
'* modDevTools provides functions for easing the process of keeping VBA code modules under version control
'* using Git. It provides functions to automate the export of a project's current state, the import of modules
'* from a version-controlled repository, as well as some fancy tools to allow checking for Reference depencies,
'* exporting plain text documentation or configuration files for the repository, etc.
'*
'* == modRegularExpressionFunctions ==
'* VBA usually comes with Microsoft VBScript Regular Expressions available as a Reference, which provides a good
'* regex implementation, but which is kind of a bear to use for short, one-off regex operations. This module
'* provides functions intended to provide something more like the easy, one-line regex functionality available in
'* scripting languages like Perl, PHP, Python, etc. (So to check for regex matching, you don't need to instantiate
'* an object, set up its properties, etc. etc.; just use RegexMatch().)
'**
Option Explicit

Public Const EX_FILEALREADYEXISTS = 58
Public Const EX_FILEPERMISSIONDENIED = 75
Public Const EX_INVALIDPROPERTYVALUE = 380
Public Const EX_DUPLICATE_KEY_VALUE = 3022

Public Const COLOR_ALERT As Long = &H10FFFF         'Bright yellow
Public Const COLOR_DISABLED As Long = &HC0C0C0      'Light grey
Public Const COLOR_UNMARKED As Long = &HFFFFFF      'White
Public Const COLOR_MARKEDERROR As Long = &HC0C0FF   'Light red

'**
'* DebugDump: Utility Function mainly for use in the Immediate pane to more easily display a
'* bunch of different kinds of objects and collections of objects in VBA
'*
'* @param Variant v The object to print out a representation of in the Immediate pane
'**
Sub DebugDump(v As Variant)
    Dim vScalar As Variant
    If IsArray(v) Or TypeName(v) = "Collection" Or TypeName(v) = "ISubMatches" Or TypeName(v) = "Fields" Then
        If TypeName(v) = "Range" Then
            Debug.Print TypeName(v.Value), LBound(v.Value), UBound(v.Value)
        ElseIf IsArray(v) Then
            Debug.Print TypeName(v), LBound(v), UBound(v)
        Else
            Debug.Print TypeName(v), v.Count
        End If
        
        For Each vScalar In v
            DebugDump (vScalar)
        Next vScalar
    ElseIf TypeName(v) = "Dictionary" Then
        Debug.Print TypeName(v)
        For Each vScalar In v.Keys
            Debug.Print vScalar & ":"
            DebugDump v.Item(vScalar)
        Next vScalar
    ElseIf TypeName(v) = "Nothing" Then
        Debug.Print TypeName(v)
    Else
        Debug.Print TypeName(v), v
    End If
End Sub

'**
'* Merge: quickly merge together the contents of any set of For Each-enumerable lists.
'* Example: Set Fish = Merge(Coll(1, 2), Coll("Red"), Coll("Blue")) will return a Collection [1, 2, "Red", "Blue"]
'*
'* @param mixed Value,... A list of lists to merge, each of which MUST be enumerable with For Each
'* @return Collection A single Collection consisting of every item in each list.
'**
Public Function Merge(ParamArray Value() As Variant) As Collection
    Dim I As Long
    Dim Item As Variant
    
    Set Merge = New Collection
    
    For I = LBound(Value) To UBound(Value)
        For Each Item In Value(I)
            Merge.Add Item
        Next Item
    Next I
End Function

'**
'* Coll: quickly construct a Collection of objects.
'* Example: Set Fish = Coll(1, 2, "Red", "Blue") will construct a Collection with the elements [1, 2, "Red", "Blue"]
'*
'* @param mixed Value,... A list of individual values to collect
'* @return Collection A collection consisting of the parameters, added in the order provided.
'**
Public Function Coll(ParamArray Value() As Variant) As Collection
    Dim I As Long
    
    Set Coll = New Collection
    For I = LBound(Value) To UBound(Value)
        Coll.Add Item:=Value(I)
    Next I
End Function

'**
'* Assoc: quickly construct a Dictionary/associative array mapping keys to values.
'* Example 1: Set Map = Assoc("METHOD", "GET", "Path", sUrlPath) will construct a Dictionary
'* with the key-value pairs {"METHOD": "GET", "Path": sUrlPath}
'* Example 2: Set Map = Assoc(Coll("METHOD", "PUT"), Coll("Path", sUrlPath)) will construct a Dictionary
'* with the key-value pairs {"METHOD": "PUT", "Path": sUrlPath}
'*
'* @param mixed KeyValue,... a list of individual values alternating between key (first) and value (second) to associate, OR a list of key-value pairs in collections
'* @return Dictionary An associative array mapping each key to the value immediately following it in the list.
'**
Public Function Assoc(ParamArray KeyValue() As Variant) As Dictionary
    Dim I As Long
    Dim Step As Integer
    Dim vKey As Variant
    Dim bIsObject As Boolean
    Dim vValue As Variant
    Dim oValue As Object
    
    Set Assoc = New Dictionary
    Let I = LBound(KeyValue)
    Do Until I > UBound(KeyValue)
        Let vKey = Null
        Let vValue = Null
        
        Let bIsObject = False
        If TypeName(KeyValue(I)) = "Collection" Then
            Let Step = 1
            Let vKey = KeyValue(I).Item(1)
            Let vValue = KeyValue(I).Item(2)
        Else
            Let Step = 2
            Let vKey = KeyValue(I)
            If I + 1 <= UBound(KeyValue) Then
                Let bIsObject = IsObject(KeyValue(I + 1))
                If bIsObject Then
                    Set oValue = KeyValue(I + 1)
                Else
                    Let vValue = KeyValue(I + 1)
                End If
            End If
        End If
        
        If bIsObject Then
            Assoc.Add Key:=vKey, Item:=oValue
        Else
            Assoc.Add Key:=vKey, Item:=vValue
        End If
        
        Let I = I + Step
    Loop
End Function

'**
'* ListOfItems: extract an enumerable list of items from many different kinds of list-y data structures
'*
'* @param Variant List
'* @return Variant
'**
Public Function ListOfItems(List As Variant) As Variant
    Dim vItem As Variant
    
    Select Case TypeName(List)
    Case "Collection":
        Set ListOfItems = List 'Collection
    Case "Dictionary":
        Set ListOfItems = New Collection
        For Each vItem In List.Items 'Variant()
            ListOfItems.Add Item:=vItem
        Next vItem
    Case Else:
        If IsObject(List) Then
            Set ListOfItems = List
        ElseIf IsArray(List) Then
            Set ListOfItems = New Collection
            For Each vItem In List 'Variant()
                ListOfItems.Add Item:=vItem
            Next vItem
        Else
            Let ListOfItems = List
        End If
    End Select
End Function

'**
'* ListSum: add together all the elements in a For Each-enumerable list and return the result
'*
'* @param Variant List
'* @param Variant Initial value to begin the summing with (if omitted, begins with 0.0)
'**
Public Function ListSum(List As Variant, Optional ByVal Initial As Variant) As Variant
    Dim vSum As Variant
    Dim vItem As Variant
    Dim vList As Variant
    
    If IsMissing(Initial) Then
        Let Initial = 0#
    End If
    
    Set vList = ListOfItems(List)
    Let vSum = Initial
    For Each vItem In vList
        Let vSum = vSum + vItem
    Next vItem
    
    Let ListSum = vSum
End Function

'**
'* BubbleSortList: sort the elements of a random-access list
'*
'* @param Variant List
'**
Public Sub BubbleSortList(ByRef List As Variant)
    Dim Swapped As Boolean
    Dim vSwap As Variant
    Dim I As Integer, J As Integer
    Dim Isub0 As Integer, IsubN As Integer
    
    If IsArray(List) Or TypeName(List) = "Collection" Then
        If IsArray(List) Then
            Isub0 = LBound(List)
            IsubN = UBound(List)
        ElseIf TypeName(List) = "Collection" Then
            Isub0 = 1
            IsubN = List.Count
        End If
            
        Do
            Let Swapped = False
            For I = Isub0 To IsubN - 1
                If List(I + 1) < List(I) Then
                    Let vSwap = List(I)
                    Let List(I) = List(I + 1)
                    Let List(I + 1) = vSwap
                    
                    Let Swapped = True
                End If
            Next I
        Loop Until Not Swapped
    End If
End Sub

'**
'* Join: join an iterable list into a string, with items separated by a given delimiter
'* (for example ["One", "Two", "Red", "Blue"] => "One;Two;Red;Blue")
'*
'* @param String Delimiter The characters used to separate items from the list (e.g.: ", ")
'* @param Variant List The items to join into a single string (can be any list iterable with For Each) (e.g. Array of ("One", "Two", "Red", "Blue"))
'* @return String containing all the items from List, separated by Delimiter (e.g. "One, Two, Red, Blue")
'**
Public Function Join(ByVal Delimiter As String, List As Variant) As String
    Dim First As Boolean
    Dim vItem As Variant
    Dim sConjunction As String
    
    First = True
    For Each vItem In List
        If Not First Then
            sConjunction = sConjunction & Delimiter
        End If
        
        sConjunction = sConjunction & vItem
        
        First = False
    Next vItem
    
    Join = sConjunction
End Function

'**
'* camelCase: convert a list of words into a CamelCase string
'* Example: "Make this into camel case" => "MakeThisIntoCamelCase"
'*
'* @param Variant Words
'* @param String FilterOut
'* @param Boolean InitialLower
'* @return String
'**
Public Function camelCase(ByVal Words As Variant, Optional ByVal FilterOut As String, Optional ByVal InitialLower As Boolean) As String
    Dim s As String
    Dim Word As Variant
    Dim Item As Variant
    Dim WordList As Variant
    
    If TypeName(Words) = "String" Then
        Let WordList = RegexSplit(Text:=Words, Pattern:="\s+")
    Else
        Set WordList = New Collection
        For Each Item In Words
            For Each Word In RegexSplit(Text:=Item, Pattern:="\s+")
                WordList.Add Word
            Next Word
        Next Item
    End If
    
    Let s = ""
    For Each Word In WordList
        If Len(FilterOut) > 0 Then
            Let Word = RegexReplace(Value:=Word, Pattern:=FilterOut, Replace:="")
        End If
        Let Word = TitleCase(Text:=Word, ForceLower:=True)
        Let s = s & Word
    Next Word
    
    If InitialLower Then
        Let s = LCase(Left(s, 1)) & Right(s, Len(s) - 1)
    End If
    
    If TypeName(WordList) = "Collection" Then
        Set WordList = Nothing
    End If
    
    Let camelCase = s
End Function

'**
'* camelCaseSplitString: split a CamelCase string into its apparent component words
'* as marked by the shifts in case
'*
'* @param String s The camelCase text to split into words
'*
'* @return Collection of String items for each word
'**
Public Function camelCaseSplitString(ByVal s As String) As Collection
    Dim isAlpha As New RegExp
    Dim isUpper As New RegExp
    Dim isLower As New RegExp
    Dim isUpperLower As New RegExp
    Dim isWhiteSpace As New RegExp
    
    With isUpper
        .IgnoreCase = False
        .Pattern = "^([A-Z])$"
    End With
    
    With isLower
        .IgnoreCase = False
        .Pattern = "^([a-z])$"
    End With
    
    With isAlpha
        .IgnoreCase = False
        .Pattern = "^([A-Za-z]+)$"
    End With

    With isUpperLower
        .IgnoreCase = False
        .Pattern = "^([A-Z][a-z])$"
    End With
    
    With isWhiteSpace
        .IgnoreCase = False
        .Pattern = "^((\s|[_])+)$"
    End With

    
    Dim cWords As New Collection
    Dim c0 As String, c As String, c2 As String
    Dim I0 As Integer, I As Integer
    Dim Anchor As Integer
    Dim State As Integer
    
    Anchor = 0
    I = 1
    GoTo NextWord
    
    'Finite State Machine
NextWord:
    If I > Len(s) Then
        GoTo ExitMachine
    End If
    
    c0 = Mid(s, I, 1)
    If isUpper.Test(c0) Then
        Anchor = I
        GoTo WordBeginsOnUpper
    ElseIf isLower.Test(c0) Then
        Anchor = I
        GoTo WordBeginsOnLower
    ElseIf isWhiteSpace.Test(c0) Then
        Anchor = I
        Let I = I + Len(c)
        GoTo NextWord
    Else
        Anchor = I
        GoTo FromOtherToNextWord
    End If

WordBeginsOnLower:
    If I > Len(s) Then GoTo NextWord
    Let c0 = c: Let c = Mid(s, I, 1)
    
    Let I = I + Len(c)
    GoTo ContinueWordToUpperBreak

ContinueWordToUpperBreak:
    If I > Len(s) Then GoTo NextWord
    Let c0 = c: Let c = Mid(s, I, 1)
    
    If isUpper.Test(c) Or isWhiteSpace.Test(c) Then
        GoTo ClipWord
    Else
        I = I + Len(c)
    End If
    GoTo ContinueWordToUpperBreak
    
WordBeginsOnUpper:
    If I > Len(s) Then GoTo NextWord
    Let c0 = c: c = Mid(s, I, 1)

    'Move ahead to the next character
    'UPPERCase: two uppers in a row
    'MixedCase: one upper, one lower
    Let I = I + 1
    Let c0 = c: c = Mid(s, I, 1)
    
    If isLower.Test(c) Then
        Let I = I + Len(c)
        GoTo ContinueWordToUpperBreak
    ElseIf isUpper.Test(c) Then
        GoTo ContinueWordToUpperLowerBreak
    ElseIf Not isWhiteSpace.Test(c) Then
        GoTo ContinueWordToUpperLowerBreak
    Else
        GoTo ClipWord
    End If
    
ContinueWordToUpperLowerBreak:
    If I > Len(s) Then GoTo NextWord
    Let c0 = c: c = Mid(s, I, 1): c2 = Mid(s, I, 2)

    If isUpperLower.Test(c2) Or isWhiteSpace.Test(c) Then
        GoTo ClipWord
    ElseIf isAlpha.Test(c) Then
        I = I + 1
        GoTo ContinueWordToUpperLowerBreak
    Else
        I = I + 1
        GoTo ContinueWordToUpperLowerBreak
    End If
    
FromOtherToNextWord:
    If I > Len(s) Then GoTo NextWord
    c = Mid(s, I, 1)
    
    If isUpper.Test(c) Or isLower.Test(c) Or isWhiteSpace.Test(c) Then
        GoTo ClipWord
    Else
        I = I + 1
        GoTo FromOtherToNextWord
    End If

ClipWord:
    If Anchor < I Then
        cWords.Add Mid(s, Anchor, I - Anchor)
    End If
    GoTo NextWord

ExitMachine:
    If Anchor > 0 Then
        cWords.Add Mid(s, Anchor, I - Anchor)
    End If
       
    Set camelCaseSplitString = cWords
End Function

'**
'* TitleCase: convert a string into Title Case (capitalized first character;
'* preserve case for the rest of the word).
'*
'* @param String s The camelCase text to split into words
'* @param Boolean ForceLower If words include uppercase letters after the first, should they be converted to lowercase?
'*
'* @return Collection of String items for each word
'**
Public Function TitleCase(ByVal Text As String, Optional ByVal ForceLower As Boolean)
    Dim Word As String
    Dim aWords() As String
    Dim I As Integer
    Dim sOutput As String
    Dim First As String, Rest As String
    
    Let aWords = RegexSplit(Text:=Text, Pattern:="\s", DelimCapture:=True)
    Let I = LBound(aWords)
    Do Until I > UBound(aWords)
        Let Word = aWords(I)
        
        Let First = Left(Word, 1)
        If First >= "a" And First <= "z" Then
            Let First = UCase(First)
        End If
        
        Let Rest = Right(Word, Len(Word) - 1)
        If ForceLower Then
            Let Rest = LCase(Rest)
        End If
        
        Let Word = First & Rest
        
        Let sOutput = sOutput & Word
        
        'Next!
        Let I = I + 1
    Loop
    
    Let TitleCase = sOutput
End Function

'**
'* FindSlugInDirectory: recursively search a directory culture for a filename slug
'*
'* @param String Directory a path to the top-level directory to search within
'* @param String Slug
'* @param Variant MaxDepth
'* @return String a path to the location of the file or folder with the slug as its name
'**
Public Function FindSlugInDirectory(Directory As String, slug As String, Optional ByVal MaxDepth As Variant) As String
    Dim sPath As String, sFoundFile As String
    Dim vSubDirectory As Variant
    Dim cSubDirectories As New Collection
    Dim FS As New FileSystemObject
    
    If IsMissing(MaxDepth) Then
        Let MaxDepth = -1
    End If
    
    If FS.FolderExists(Directory & "\" & slug) Or FS.FileExists(Directory & "\" & slug) Then
        Let FindSlugInDirectory = Directory & "\" & slug
    Else
        Let sPath = Dir(PathName:=Directory & "\*.*", Attributes:=vbDirectory)
        Do Until Len(sPath) = 0
            If Not RegexMatch(sPath, "^[.]+$") Then
                If FS.FolderExists(Directory & "\" & sPath) Then
                    cSubDirectories.Add Item:=sPath
                End If
            End If
            Let sPath = Dir(Attributes:=vbDirectory)
        Loop
        
        For Each vSubDirectory In cSubDirectories
            Let sPath = CStr(vSubDirectory)
            If MaxDepth <> 0 Then
                Let sFoundFile = FindSlugInDirectory(Directory:=Directory & "\" & sPath, slug:=slug, MaxDepth:=MaxDepth - 1)
                If Len(sFoundFile) > 0 Then
                    Exit For
                End If
            End If
        Next vSubDirectory
        
        Let FindSlugInDirectory = sFoundFile
    End If
End Function

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

'**
'* GetOption: get the value of a persistent setting stored in a database alongside the project
'*
'* @param String SettingName name of the option or options to retrieve
'*    This can be a list of settings separated by pipes (e.g.: "VBA.Repository.MyVBA|VBA.Repository")
'*    Given more than setting name, GetOption will check the first key (e.g.: "VBA.Repository.MyVBA") for a value;
'*    if the first key has no value associated with it, then it will check the second key (e.g.: "VBA.Repository");
'*    and so on. If no key has values associated with it, then the Default value will be returned.
'* @param Variant Default value to return if the setting does not exist.
'* @return Variant Value set for the option.
'**
Public Function GetOption(SettingName As String, Optional Default As Variant) As Variant
    Dim v As Variant
    Dim Qdf As DAO.QueryDef
    Dim Rs As DAO.Recordset
    
    Dim I As Integer
    Dim N As Integer
    Dim sSettingName As String
    Dim asSettingName() As String
    
    Let asSettingName = Split(Expression:=SettingName, Delimiter:="|")
    If UBound(asSettingName) >= LBound(asSettingName) Then
        
        Let v = Empty
        Let N = UBound(asSettingName)
        
        For I = LBound(asSettingName) To N
            Let sSettingName = asSettingName(I)
            
            Set Qdf = CurrentDb.CreateQueryDef(Name:="", SQLText:="SELECT * FROM CommonSettings WHERE SettingName=[paramSettingName]")
            Let Qdf.Parameters("paramSettingName") = sSettingName
    
            Set Rs = Qdf.OpenRecordset
    
            If Not Rs.EOF Then
                Let v = Rs!SettingValue.Value
            End If
            
            Rs.Close
            Set Rs = Nothing
            Set Qdf = Nothing
            
            If Not IsEmpty(v) Then
                Exit For
            End If
        Next I
    
        If IsEmpty(v) Then
            'No such setting in the table
            If Not IsMissing(Default) Then
                Let v = Default
            End If
            
        End If
        
        Let GetOption = v
    End If
End Function


