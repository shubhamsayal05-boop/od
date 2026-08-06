Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
' Cellule Office

Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.UseSystemSeparators = True
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

End Sub
Sub wProtect()
    
    Dim timeStamp As String
    
    If ThisWorkbook.ReadOnly = False Then
        If Not CBool(Application.UserName Like "*E592622*") = True Then
'            ThisWorkbook.Close False
             timeStamp = Format(Now, "yyyymmdd_hhmmss")
             ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\sav_" & timeStamp & ".xlsm"
            Exit Sub
        End If
     End If
    If Not CBool(Application.UserName Like "*E592622*") = True Then
        Application.ShowDevTools = False
        Application.OnKey "%{F11}", " "
        On Error Resume Next
        ThisWorkbook.VBProject.VBE.MainWindow.Visible = False
        On Error GoTo 0
    End If
    
End Sub

Private Sub Workbook_Open()

    If Not ValidateMainDatabase() Then
        MsgBox "DriveScope database not found." & vbCrLf & vbCrLf & _
               "Expected file:" & vbCrLf & GetMainDbPath() & vbCrLf & vbCrLf & _
               "Run scripts/setup_drivescope.py or copy _OdrivDB.accdb into data/db.", vbCritical, "DriveScope"
        Exit Sub
    End If

    Call FiltresOff
    Call HideFeuille(2)
    Call DBStructure.getEntete
    Call DBStructure.getInterval
    db.Terminate
   
    ThisWorkbook.sheets("HOME").Range("Version").Value = "VERSION" & " " & Left(Mid(ThisWorkbook.Name, 7), Len(Mid(ThisWorkbook.Name, 7)) - 5)
    ThisWorkbook.sheets("HOME").Activate

    Call checkDevVersion("1.0.0")
    Application.UseSystemSeparators = False
    Application.DecimalSeparator = "."
    Application.ThousandsSeparator = " "
    AutoBackupWorkbook
End Sub
Private Sub Workbook_SheetActivate(ByVal sh As Object)
    HideFeuille xlSheetVeryHidden
End Sub

Function HideFeuille(options As Integer)
   Dim tables(13) As String
   Dim i As Integer
    
    tables(1) = "CONFIGURATIONS SEETINGS"
    tables(2) = "SETTINGS"
    tables(3) = "structure"
    tables(4) = "POWERTRAIN"
    tables(5) = "CONFIGURATIONS"
    tables(6) = "TARGET VEHICLE"
    tables(7) = "TARGETS"
    tables(8) = "Calculs"
    tables(9) = "Graph_status"
    tables(10) = "DEFINITION SDV"
    tables(11) = "PARAMETRES GRAPH"
    tables(12) = "ENTETE_COLONNE"
    tables(13) = "SDV MANAGER"
    
    For i = 1 To UBound(tables)
             If sheets(tables(i)).Visible <> options Then
                   sheets(tables(i)).Visible = 2
             End If
    Next i
'
   
End Function





