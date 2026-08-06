Attribute VB_Name = "F_Home"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Sub Erase_All()                                        '[Effacer toutes les données]
    On Error GoTo gEr
    If ThisWorkbook.sheets("HOME").Range("Fuel").Value = "" Or ThisWorkbook.sheets("HOME").Range("Gears").Value = "" Or ThisWorkbook.sheets("HOME").Range("Software") = "" Or ThisWorkbook.sheets("HOME").Range("Prestation").Value = "" Or ThisWorkbook.sheets("HOME").Range("DriveVersion").Value = "" Or ThisWorkbook.sheets("HOME").Range("Milestone").Value = "" Or ThisWorkbook.sheets("HOME").Range("Area").Value = "" Then
        MsgBox "No project to Remove.", vbCritical
    Else
        EventAndScreen (False)
        Call DeleteProject
        ThisWorkbook.sheets("Graph_status").Range("I2") = ""
        ThisWorkbook.sheets("Graph_status").Range("I20") = ""
        EventAndScreen (True)
    End If

gEr:
    If ERR.Number <> 0 Then
        EventAndScreen (True)
        MsgBox ERR.description, vbCritical, "ODRIV"
    End If
End Sub

Sub DeleteProject()
    
    If MsgBox("Are you sure?" & vbCr & "This operation will delete all your data.", vbOKCancel Or vbQuestion, "Please confirm") = vbOK Then
        Application.StatusBar = "Erasing all datas"
        Application.ScreenUpdating = False
        ThisWorkbook.sheets("structure").Range("N1") = 2
        Erase_All2
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "Done", vbInformation
    End If

End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$H$10" Or Target.Address = "$H$16" Then
        If Len(ThisWorkbook.sheets("home").Range("Fuel").Value) > 0 And Len(ThisWorkbook.sheets("home").Range("DriveVersion").Value) > 0 Then
            MajTargets
        End If
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'If target.Address = "$C$8:$K$8" Then
    '    If Application.ScreenUpdating Then frmNameCode.Show
    'End If
End Sub


