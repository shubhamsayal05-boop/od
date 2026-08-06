Attribute VB_Name = "ClasseBdD"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private oDB As Object
Private oDB1 As Object
Private oDB2 As Object
Private oDB3 As Object
Private oDB4 As Object
Private curr_ip As String
Private alive As Boolean

Private Sub Class_Initialize()
    
    Dim timeStamp As String
    
    Application.StatusBar = "Init DB in progress..."
    
    Set oDB = CreateObject("ADODB.Connection")
    Set oDB1 = CreateObject("ADODB.Connection")
    Set oDB2 = CreateObject("ADODB.Connection")
    Set oDB3 = CreateObject("ADODB.Connection")
    Set oDB4 = CreateObject("ADODB.Connection")
    
    On Error GoTo gErr
    
    
    
    oDB.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & GetMainDbPath() & ";"
    oDB.Open
    
    
    On Error GoTo 0
    Application.StatusBar = False
    curr_ip = GetIPAddress
    alive = True
    Exit Sub

gErr:
    Die "Erreur accès Base de données, contactez la cellule support." '' & Chr(10) & ERR.description
'    MsgBox "Fermeture d'Odriv", vbCritical, "ODRIV"
'    ThisWorkbook.Close False
     ERR.Clear
     timeStamp = Format(Now, "yyyymmdd_hhmmss")
     ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\sav_" & timeStamp & ".xlsm"
End Sub

Private Function openConnection(ByRef db As Object, ByVal dbName As String)
    
    Dim timeStamp As String
    
    On Error GoTo ERR
    db.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & GetYearDbPath(dbName) & ";"
    db.Open
    If db.Errors.Count > 0 And db.state <> 1 Then Die "Erreur lors de la connexion à la base de données, " & dbName
    Exit Function
ERR:
    Die "Erreur Accès Base de données " & dbName & ERR.description
    ERR.Clear
    timeStamp = Format(Now, "yyyymmdd_hhmmss")
    ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\sav_" & timeStamp & ".xlsm"
'    MsgBox "Fermeture d'Odriv", vbCritical, "ODRIV"
'    ThisWorkbook.Close False
End Function
Private Sub Class_Terminate()
    closeConnection oDB
    
    closeConnection oDB1
    closeConnection oDB2
    closeConnection oDB3
    closeConnection oDB4
    
    DoEvents
End Sub

Public Sub Terminate()
    Class_Terminate
End Sub
Private Function closeConnection(ByRef db As Object)
    If Not db Is Nothing Then
        If db.state Then db.Close
''''        Set db = Nothing
    End If
End Function

Public Property Get Status() As Boolean
    Dim ip As String
    ip = GetIPAddress
    If oDB.Errors.Count = 0 Then If alive Then Status = CBool(oDB.state = 1)

End Property

Function Request(ByVal sql As String, Optional db As Object) As Object
    Dim RS As Object
    Dim nbMax As Integer
    Dim timeStamp As String
    
    Const adOpenStatic = 2
    
    Set RS = CreateObject("ADODB.Recordset")
    RS.cursortype = 0
    RS.cursorlocation = 3
    RS.stayinsync = False


    On Error GoTo gErr
        If Not db Is Nothing Then
            RS.Open sql, db, adOpenStatic, 1
        Else
             RS.Open sql, oDB, adOpenStatic, 1
        End If
    On Error GoTo 0

    If Not RS.EOF Then
        RS.movefirst
        Set Request = RS
        Set RS = Nothing
    End If
    jLog sql
    Exit Function

gErr:
    jLog "[E]" & sql & " : " & ERR.description
    MsgBox "Erreur base de donnée " & vbCrLf & ERR.description, vbCritical, "ODRIV"
    ERR.Clear
    timeStamp = Format(Now, "yyyymmdd_hhmmss")
    ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\sav_" & timeStamp & ".xlsm"
'    MsgBox "Fermeture d'Odriv", vbCritical, "ODRIV"
'    ThisWorkbook.Close False
    
    'If Err.Number = -2147467259 Then
    '    Set oDB = Nothing
    '    Call Class_Initialize
    '    nbMax = nbMax + 1
    '    If nbMax < 10 Then Resume
    'End If

End Function

Function GetValue(ByVal sql As String, Optional db As Object) As Variant
    Dim RS As Object
    Dim timeStamp As String

    On Error GoTo gErr
    Set RS = CreateObject("ADODB.Recordset")
    RS.cursortype = 0
    RS.cursorlocation = 3
    RS.stayinsync = False

   
    If Not db Is Nothing Then
         RS.Open sql, db, 2, 1
    Else
          RS.Open sql, oDB, 2, 1
    End If
    
    On Error GoTo gErr
    If Not RS.EOF And Not IsNull(RS.Fields(0)) Then GetValue = RS.Fields(0).Value
    RS.Close
    Set RS = Nothing
    Exit Function

gErr:
    jLog "[E]" & sql & " : " & ERR.description
    MsgBox "Erreur base de donnée " & vbCrLf & ERR.description, vbCritical, "ODRIV"
    ERR.Clear
    timeStamp = Format(Now, "yyyymmdd_hhmmss")
    ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\sav_" & timeStamp & ".xlsm"
'    MsgBox "Fermeture d'Odriv", vbCritical, "ODRIV"
'    ThisWorkbook.Close False
End Function

Function GetListValue(ByVal sql As String, Optional db As Object) As Variant
    Dim RS As Object
    Dim timeStamp As String
    
    On Error GoTo gErr
    Set RS = CreateObject("ADODB.Recordset")
    RS.cursortype = 0
    RS.cursorlocation = 3
    RS.stayinsync = False

   
    If Not db Is Nothing Then
         RS.Open sql, db, 2, 1
    Else
          RS.Open sql, oDB, 2, 1
    End If
    
    On Error GoTo gErr
    If Not RS.EOF And Not IsNull(RS.Fields(0)) Then
    GetListValue = RS.getrows
     jLog sql
    End If
     RS.Close
    Set RS = Nothing
    Exit Function

gErr:
    jLog "[E]" & sql & " : " & ERR.description
    MsgBox "Erreur base de donnée " & vbCrLf & ERR.description, vbCritical, "ODRIV"
    ERR.Clear
    timeStamp = Format(Now, "yyyymmdd_hhmmss")
    ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\sav_" & timeStamp & ".xlsm"
'    MsgBox "Fermeture d'Odriv", vbCritical, "ODRIV"
'    ThisWorkbook.Close False
End Function
Sub Execute(ByVal sql As String, Optional db As Object)
  
    Dim timeStamp As String
    
    On Error GoTo gErr
    
    If Not db Is Nothing Then
         db.Execute sql
    Else
         oDB.Execute sql
    End If
    
    
    jLog "Exec " & sql
    Exit Sub
gErr:
        If ERR.Number <> 0 And ERR.Number <> -2147168234 Then
            jLog "[E]" & sql & " : " & ERR.description
        End If
        If ERR.Number <> 0 Then
            MsgBox "Erreur base de donnée " & vbCrLf & ERR.description, vbCritical, "ODRIV"
            ERR.Clear
            timeStamp = Format(Now, "yyyymmdd_hhmmss")
            ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\sav_" & timeStamp & ".xlsm"
'            MsgBox "Fermeture d'Odriv", vbCritical, "ODRIV"
'            ThisWorkbook.Close False
        End If
End Sub

Function GetIPAddress() As String
    Static objWMIService As Object
    Dim IPConfigSet As Object, IPConfig As Object
    Dim strIPAddress As String, IPAddress

    ' Connect to the WMI service
    If objWMIService Is Nothing Then Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

    ' Get all TCP/IP-enabled network adapters
    Set IPConfigSet = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")

    ' Get all IP addresses associated with these adapters
    For Each IPConfig In IPConfigSet
        IPAddress = IPConfig.IPAddress
        If Not IsNull(IPAddress) Then
            strIPAddress = strIPAddress & Join(IPAddress, ", ")
        End If
    Next
    Set IPConfigSet = Nothing
    GetIPAddress = strIPAddress
End Function

Function GetOdb(Optional db As Integer) As Object
    If oDB Is Nothing Then Call Class_Initialize
     If db <> 0 Then
            If db = 1 Then
'                If oDB1.connectionString = "" Or oDB1 Is Nothing Then openConnection oDB1, "_OdrivDB_1"
                CloseSudbConn
                openConnection oDB1, "_OdrivDB_1"
                Set GetOdb = oDB1
            End If
            
             If db = 2 Then
'                If oDB2.connectionString = "" Or oDB2 Is Nothing Then openConnection oDB2, "_OdrivDB_2"
                CloseSudbConn
                openConnection oDB2, "_OdrivDB_2"
                Set GetOdb = oDB2
            End If
            
             If db = 3 Then
'                If oDB3.connectionString = "" Or oDB3 Is Nothing Then openConnection oDB3, "_OdrivDB_3"
                 CloseSudbConn
                 openConnection oDB3, "_OdrivDB_3"
                Set GetOdb = oDB3
            End If
            
             If db = 4 Then
'                If oDB4.connectionString = "" Or oDB4 Is Nothing Then openConnection oDB4, "_OdrivDB_4"
                CloseSudbConn
                openConnection oDB4, "_OdrivDB_4"
                Set GetOdb = oDB4
            End If
    Else
        Set GetOdb = oDB
    End If
End Function
Function AnneeEnCours() As String
     AnneeEnCours = ThisWorkbook.sheets("CFG").Range("B3")
End Function

Function GetConnFromDbName(dbName As String) As Object
     If Not dbName = "_OdrivDB_1" Then Set GetConnFromDbName = oDB1
     If Not dbName = "_OdrivDB_2" Then Set GetConnFromDbName = oDB2
     If Not dbName = "_OdrivDB_3" Then Set GetConnFromDbName = oDB3
     If Not dbName = "_OdrivDB_4" Then Set GetConnFromDbName = oDB4
  
End Function

Function CloseSudbConn()
    closeConnection oDB1
    closeConnection oDB2
    closeConnection oDB3
    closeConnection oDB4
End Function




