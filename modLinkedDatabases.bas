Attribute VB_Name = "modLinkedDatabases"
Option Compare Database
Option Explicit

Function IsLinkedToDatabase(Table As DAO.TableDef, DatabasePath As String, Optional ByRef DB As DAO.Database)
    If IsMissing(DB) Or DB Is Nothing Then
        Set DB = CurrentDb
    End If
    
    IsLinkedToDatabase = (Table.Connect Like "*;DATABASE=" & DatabasePath & "*")
End Function

Function IsAdjacentToAllLinkedDatabases(Optional ByRef DB As DAO.Database) As Boolean
    Dim sMyPath As String
    Dim sDbPath As String
    Dim sFromDB As String
    
    Dim bNotAdjacent As Boolean
    
    Dim oTableLink As DAO.TableDef

    Dim FSO As New FileSystemObject
    
    If IsMissing(DB) Or DB Is Nothing Then
        Set DB = CurrentDb
    End If
    
    Let sMyPath = CurrentProject.Path
    
    Let bNotAdjacent = False
    
    For Each oTableLink In DB.TableDefs
        If Len(oTableLink.SourceTableName) > 0 Then
            Let sFromDB = RegexReplace(Value:=oTableLink.Connect, Pattern:="^.*DATABASE=([^;]+)(;.*)?$", Replace:="$1")
            If IsLinkedToDatabase(Table:=oTableLink, DatabasePath:=sFromDB, DB:=DB) Then
                Let sDbPath = FSO.GetParentFolderName(sFromDB)
                If sDbPath <> sMyPath Then
                    Let bNotAdjacent = True
                End If
            End If
        End If
    Next oTableLink
    
    Let IsAdjacentToAllLinkedDatabases = Not (bNotAdjacent)
End Function

Function HasAdjacentDatabaseCopies(Optional ByRef DB As DAO.Database) As Boolean
    Dim sMyPath As String
    Dim sDbFile As String
    Dim sDbPath As String
    Dim sFromDB As String
    Dim sToDB As String
    
    Dim bAllAdjacent As Boolean
    
    Dim oTableLink As DAO.TableDef

    Dim FSO As New FileSystemObject
    
    If IsMissing(DB) Or DB Is Nothing Then
        Set DB = CurrentDb
    End If
    
    Let sMyPath = CurrentProject.Path
    
    Let bAllAdjacent = True
    
    For Each oTableLink In DB.TableDefs
        If Len(oTableLink.SourceTableName) > 0 Then
            Let sFromDB = RegexReplace(Value:=oTableLink.Connect, Pattern:="^.*DATABASE=([^;]+)(;.*)?$", Replace:="$1")
            If IsLinkedToDatabase(Table:=oTableLink, DatabasePath:=sFromDB, DB:=DB) Then
                Let sDbPath = FSO.GetParentFolderName(sFromDB)
                If sDbPath <> sMyPath Then
                    Let sDbFile = FSO.GetFileName(sFromDB)
                    Let sToDB = sMyPath & "\" & sDbFile
                    
                    If Not FSO.FileExists(sToDB) Then
                        Let bAllAdjacent = False
                    End If
                End If
            End If
        End If
    Next oTableLink

    Let HasAdjacentDatabaseCopies = bAllAdjacent
End Function

Sub DoReLinkDatabaseTables(FromDB As String, ToDB As String, Optional ByRef DB As DAO.Database)
    If IsMissing(DB) Or DB Is Nothing Then
        Set DB = CurrentDb
    End If
    
    Dim oTableLink As DAO.TableDef
    Dim sConnectedDB As String
    
    For Each oTableLink In DB.TableDefs
        If Len(oTableLink.SourceTableName) > 0 Then
            If IsLinkedToDatabase(Table:=oTableLink, DatabasePath:=FromDB, DB:=DB) Then
                Let sConnectedDB = Replace(Expression:=oTableLink.Connect, Find:="DATABASE=" & FromDB, Replace:="DATABASE=" & ToDB)
                If oTableLink.Connect <> sConnectedDB Then
                    Let oTableLink.Connect = sConnectedDB
                    oTableLink.RefreshLink
                End If
            End If
        End If
    Next oTableLink
    
End Sub

Sub DoReLinkDatabaseTablesLocally(Optional ByRef DB As DAO.Database)
    If IsMissing(DB) Or DB Is Nothing Then
        Set DB = CurrentDb
    End If
    
    Dim sPath As String
    Dim sConnectedDB As String
    Dim sLocalDB As String
    Dim oTableLink As DAO.TableDef
    Dim FSO As New FileSystemObject
    
    Let sPath = CurrentProject.Path
    
    For Each oTableLink In DB.TableDefs
        If Len(oTableLink.SourceTableName) > 0 Then
            If Len(oTableLink.Connect) > 0 Then
                Let sConnectedDB = RegexReplace(Value:=oTableLink.Connect, Pattern:="^.*DATABASE=([^;]+)(;.*)?$", Replace:="$1")
                Let sLocalDB = sPath & "\" & FSO.GetFileName(Path:=sConnectedDB)
                If FSO.FileExists(FileSpec:=sLocalDB) Then
                    If sConnectedDB <> sLocalDB Then
                        DoReLinkDatabaseTables FromDB:=sConnectedDB, ToDB:=sLocalDB, DB:=DB
                    End If
                End If
            End If
        End If
    Next oTableLink
    
End Sub

Function GetNonAdjacentDatabasePaths(Optional ByRef DB As DAO.Database) As String
    Dim oTableLink As DAO.TableDef
    Dim FSO As New FileSystemObject
    Dim sFromDB As String
    Dim sMyPath As String
    Dim sDbPath As String
    Dim cPaths As New Collection
    
    If IsMissing(DB) Or DB Is Nothing Then
        Set DB = CurrentDb
    End If
    
    Let sMyPath = CurrentProject.Path
    
    For Each oTableLink In DB.TableDefs
        If Len(oTableLink.SourceTableName) > 0 Then
            Let sFromDB = RegexReplace(Value:=oTableLink.Connect, Pattern:="^.*DATABASE=([^;]+)(;.*)?$", Replace:="$1")
            If IsLinkedToDatabase(Table:=oTableLink, DatabasePath:=sFromDB, DB:=DB) Then
                Let sDbPath = FSO.GetParentFolderName(sFromDB)
                If sDbPath <> sMyPath Then
                    On Error Resume Next
                    cPaths.Add Item:=sDbPath, Key:=sDbPath
                    On Error GoTo 0
                End If
            End If
        End If
    Next oTableLink

    Let GetNonAdjacentDatabasePaths = Join("; ", List:=cPaths)
    
    Set cPaths = Nothing
    
End Function

Function DoCheckForMaybeReLinkDatabases()
    
    If Not IsAdjacentToAllLinkedDatabases Then
        If HasAdjacentDatabaseCopies Then
            Dim sMyPath As String
            Dim sLinkedPaths As String
            
            
            Let sMyPath = CurrentProject.Path
            If MsgBox(Prompt:="Linked databases are in " & GetNonAdjacentDatabasePaths & ". Do you want to localize to copies in " & sMyPath, Buttons:=vbYesNo, Title:="Local Database Copies") = vbYes Then
                DoReLinkDatabaseTablesLocally
            End If
        End If
    End If
    
End Function
