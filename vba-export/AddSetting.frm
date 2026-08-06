Attribute VB_Name = "AddSetting"
Attribute VB_Base = "0{A507B821-259E-47C7-9A25-0558ED9989A3}{E30EBF35-9109-43E4-BB5F-B6A1C63354F5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Sub CommandButton1_Click()
            If Len(Me.TextBox2) = 0 Then
                    MsgBox "Aucune Valeur Saisie", vbCritical, "ODRIV"
            ElseIf IsCreate = True Then
                    MsgBox "Cette SDV Existe Déjà", vbCritical, "ODRIV"
            Else
                     CreateNew.NewSDVConfigurationSetting (Me.TextBox2)
                     MsgBox "Paramètres Ajoutés", vbInformation, "ODRIV"
                     Unload Me
            End If
End Sub
Function IsCreate() As Boolean
    Dim v
    Dim i As Integer
    IsCreate = False
    
   v = ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").UsedRange.Columns(1).Value
    For i = 1 To UBound(v, 1)
        If Len(v(i, 1)) > 0 And CStr(v(i, 1)) = Me.TextBox2 Then
            IsCreate = True
            Exit Function
        End If
    Next i
    Erase v
End Function
Private Sub CommandButton3_Click()
          If Len(Me.ComboBox3) = 0 Then
                    MsgBox "Choisir SDV", vbCritical, "ODRIV"
            Else
                     ConfigSetting.TextBox2 = Me.ComboBox3
                     ConfigSetting.TextBox2.Locked = True
                     Unload Me
                     ConfigSetting.Show
            End If
End Sub
Private Sub UserForm_Initialize()
        Me.MultiPage1.Value = 0
        Dim r As Long
        Dim i As Integer
        Application.EnableEvents = False
         ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Outline.ShowLevels RowLevels:=2
        With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
                r = .Range("A65000").End(xlUp).row
                 For i = 3 To r
                            If Len(.Cells(i, 1)) > 0 Then
                               Me.ComboBox3.AddItem .Cells(i, 1)
                            End If
                 Next i
        End With
        ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Outline.ShowLevels RowLevels:=1
        
        Application.EnableEvents = True
End Sub

