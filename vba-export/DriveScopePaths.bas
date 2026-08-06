Attribute VB_Name = "DriveScopePaths"
Option Explicit

Private Const LOCAL_DB_FOLDER As String = "data\db"
Private Const MAIN_DB_FILE As String = "_OdrivDB.accdb"
Private Const DEV_VERSION_FILE As String = "ODRIV_DEV_VERSION.txt"

Public Function GetDbRootPath() As String
    Dim configuredPath As String
    Dim localPath As String
    Dim resolvedPath As String

    configuredPath = Trim$(CStr(ThisWorkbook.Worksheets("CFG").Range("B1").Value))
    localPath = BuildLocalDbRootPath()

    If Len(configuredPath) > 0 Then
        If Left$(configuredPath, 2) = "\\" Then
            resolvedPath = localPath
        ElseIf Dir(configuredPath, vbDirectory) <> "" Then
            resolvedPath = configuredPath
        Else
            resolvedPath = localPath
        End If
    Else
        resolvedPath = localPath
    End If

    EnsureDriveScopeDataFolders resolvedPath
    SyncCfgPaths resolvedPath
    GetDbRootPath = resolvedPath
End Function

Public Function GetMainDbPath() As String
    GetMainDbPath = GetDbRootPath() & "\" & MAIN_DB_FILE
End Function

Public Function GetYearDbPath(ByVal dbName As String) As String
    GetYearDbPath = GetDbRootPath() & "\" & CStr(ThisWorkbook.Worksheets("CFG").Range("B3").Value) & "\" & dbName & ".accdb"
End Function

Public Function GetLogRootPath() As String
    GetLogRootPath = GetDbRootPath()
End Function

Public Function GetDevVersionFilePath() As String
    GetDevVersionFilePath = GetLogRootPath() & "\" & DEV_VERSION_FILE
End Function

Public Sub EnsureDriveScopeDataFolders(Optional ByVal dbRoot As String = "")
    Dim Fso As Object
    Dim yearFolder As String

    If Len(dbRoot) = 0 Then dbRoot = BuildLocalDbRootPath()

    Set Fso = CreateObject("Scripting.FileSystemObject")
    If Not Fso.FolderExists(dbRoot) Then Fso.CreateFolder dbRoot

    yearFolder = dbRoot & "\" & CStr(ThisWorkbook.Worksheets("CFG").Range("B3").Value)
    If Not Fso.FolderExists(yearFolder) Then Fso.CreateFolder yearFolder

    EnsureDevVersionFile dbRoot
End Sub

Private Function BuildLocalDbRootPath() As String
    BuildLocalDbRootPath = ThisWorkbook.Path & "\" & LOCAL_DB_FOLDER
End Function

Private Sub SyncCfgPaths(ByVal dbRoot As String)
    With ThisWorkbook.Worksheets("CFG")
        If StrComp(Trim$(CStr(.Range("B1").Value)), dbRoot, vbTextCompare) <> 0 Then
            .Range("B1").Value = dbRoot
        End If
        If StrComp(Trim$(CStr(.Range("B2").Value)), dbRoot, vbTextCompare) <> 0 Then
            .Range("B2").Value = dbRoot
        End If
    End With
End Sub

Private Sub EnsureDevVersionFile(ByVal dbRoot As String)
    Dim Fso As Object
    Dim versionFile As Object
    Dim versionPath As String

    versionPath = dbRoot & "\" & DEV_VERSION_FILE
    If Dir(versionPath) <> "" Then Exit Sub

    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set versionFile = Fso.CreateTextFile(versionPath, True)
    versionFile.WriteLine "VERSION 1.0.0"
    versionFile.Close
End Sub

Public Function ValidateMainDatabase() As Boolean
    Dim dbPath As String

    dbPath = GetMainDbPath()
    ValidateMainDatabase = (Dir(dbPath) <> "")
End Function
