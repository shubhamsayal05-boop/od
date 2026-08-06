Attribute VB_Name = "ModuleUtils"
Option Explicit

'Const fLog As String = "\\besn01\CT_VAL_MAP_BDD\BdD\CDT\CDT_Activity_EE.log"
'Const fLog As String = ThisWorkbook.Worksheets("HOME").Range("AP7").Value & "\base_ODRIV.log"
Public Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hwnd As LongPtr, ByVal lpRect As LongPtr, ByVal bErase As Boolean) As Long
Public noLog As Boolean

' BdD
Function db() As ClasseBdD
    Static ldb As ClasseBdD
    
    If ldb Is Nothing Then Set ldb = New ClasseBdD
    If Not ldb.Status Then Set ldb = New ClasseBdD

    Set db = ldb
End Function

Sub Die(ByVal msg As String)
    Dim Titre
    Dim timeStamp As String
    
    
    MsgBox msg, vbCritical, "DriveScope"
    jLog msg & " : " & Environ("computername") & ";" & Environ("username") & " die.."
'    ThisWorkbook.Close False
     timeStamp = Format(Now, "yyyymmdd_hhmmss")
     ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\sav_" & timeStamp & ".xlsm"
End Sub

Sub jLog(ByVal msg As String)
    Static Fso As Object
    Dim f As Object
    Static mem As String
    Dim fLog As String
    
    Dim st As String
    st = GetLogRootPath()
'    st = Replace(Replace(Replace(st, vbLf, ""), vbCr, ""), vbNewLine, "")
    
    fLog = Trim(st) & "\ODRIV.log"

    If StrComp(mem, msg, vbTextCompare) = 0 Then Exit Sub
    If Not noLog Then
        If InStr(1, "Erreur : " & Chr(10) & msg, "[E]", vbTextCompare) > 0 Then MsgBox msg, vbCritical, Application.UserName

        On Error Resume Next
        If Fso Is Nothing Then Set Fso = CreateObject("scripting.FileSystemObject")
        Set f = Fso.opentextfile(fLog, 8, True)
        f.Write Date & ";" & Time & ";" & getToolVersion(ThisWorkbook.Name) & ";" & ThisWorkbook.sheets("cfg").Range("L2").Value & ";" & Application.UserName & ";" & msg & vbNewLine
        f.Close

        mem = msg
    End If
    Set f = Nothing
End Sub
Function getToolVersion(tool_name) As String
    
    getToolVersion = Split(tool_name, ".")(0)
End Function


' en SQL, certains caractères sont interdits.
Function toSQL(ByVal Item As String) As String
    Item = LTrim(RTrim(Item))
    toSQL = replace(Item, "'", "''")
End Function


Function TotEventSheet(SDV As String)
            Dim i As Long
            i = 13
            TotEventSheet = 0
            With ThisWorkbook.sheets(SDV)
                While Len(.Cells(6, i).Value) > 0
                    If .Cells(.Rows.Count, i).End(xlUp).row > TotEventSheet Then
                        TotEventSheet = .Cells(.Rows.Count, i).End(xlUp).row
                    End If
                    i = i + 1
                Wend
            End With
            
End Function
    


