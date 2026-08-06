Attribute VB_Name = "defineTemp"
Attribute VB_Base = "0{DA443425-9802-49C5-97E8-010990C4D6FA}{F28BA8D2-789A-492C-B196-7ABEB500AEAD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub CommandButton1_Click()
   If val(Me.tEMPS) > 0 Then
     Call Load_Data.MkTemp(val(Me.tEMPS))
     Unload Me
   Else
      MsgBox "Entrer une valeur correcte", vbCritical, "ODRIV"
   End If
End Sub
