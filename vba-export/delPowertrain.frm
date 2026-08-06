Attribute VB_Name = "delPowertrain"
Attribute VB_Base = "0{3C0D9CC5-D478-4498-BF33-7C21BF2AB3FA}{81F65B3A-6F24-44A0-98B2-92D7631E89DE}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Function chargeVal()
    Dim v
    Dim i As Integer
    v = ThisWorkbook.sheets("POWERTRAIN").UsedRange.Value
    For i = 3 To UBound(v, 1)
        If v(i, 1) = "Titre config" Then
                Me.ComboBox3.AddItem v(i, 2)
        End If
    Next i
    
    Erase v
End Function
Private Sub CommandButton3_Click()
        If Len(Me.ComboBox3.Value) > 0 Then
            Call Dels
            MsgBox "Suppression Réussie", vbInformation, "ODRIV"
            Unload Me
        Else
            MsgBox "REMPLIR ", vbCritical, "ODRIV"
        End If
End Sub

Private Sub UserForm_Initialize()
        Call chargeVal
End Sub

Function Dels()
    Dim v
    Dim i As Integer
    Dim found As Long
    
    found = 0
    Application.EnableEvents = False
    v = ThisWorkbook.sheets("POWERTRAIN").UsedRange.Value
    For i = 3 To UBound(v, 1)
        If v(i, 1) = "Titre config" And UCase(CStr(v(i, 2))) = UCase(Me.ComboBox3) Then found = i
        If UCase(v(i, 1)) = "SOMME" And found <> 0 Then
                ThisWorkbook.sheets("POWERTRAIN").Rows(found & ":" & i).EntireRow.Delete
'                 ThisWorkbook.Sheets("POWERTRAIN").Rows(found & ":" & i).Select
                 Application.EnableEvents = True
                Exit Function
        End If
    Next i
    Application.EnableEvents = True
    Erase v
End Function
