Attribute VB_Name = "frmNameCode"
Attribute VB_Base = "0{09B5512B-9813-4D45-9750-BDFA0AEEBBCC}{3440E653-397A-4D1B-AA44-AF22BB0EDE6A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Private Sub CommandButton1_Click()

    If Trim(Me.TextBox1.text) = "" Or Trim(Me.TextBox2.text) = "" Or Trim(Me.TextBox3.text) = "" Then
        MsgBox "Tous les champs sont nécessaires.", vbCritical, "ODRIV"
    Else
       
        Me.hide
        ThisWorkbook.sheets("home").Range("Project").Value = Me.TextBox1.text & "_" & Me.TextBox2.text & "_" & Me.TextBox3.text
        'db.Execute "UPDATE projet SET name='" & ThisWorkbook.Sheets("HOME").Range("project").Value & "' WHERE id= '" & id_projet & "' "
        
        ThisWorkbook.sheets("home").Range("Project").Offset(1, 0).Select
    End If
End Sub

Private Sub UserForm_Activate()
    Dim INFO() As String

    ThisWorkbook.sheets("home").Range("Project").Value = replace(ThisWorkbook.sheets("home").Range("Project").Value, " ", "_")
    INFO = Split(ThisWorkbook.sheets("home").Range("Project").Value, "_")
    If UBound(INFO) > -1 Then
        Me.TextBox1.text = INFO(0)
        Me.TextBox2.text = INFO(1)
        Me.TextBox3.text = INFO(2)
    End If
End Sub


