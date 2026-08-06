Attribute VB_Name = "F_Rating"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Private Sub Worksheet_Activate()
    Dim RS As Double
'    RS = GetNoteGlobalTarget("driv")
'    ActiveSheet.Range("EQ3").Value = IIf(RS = -555, "", RS)
'
'    RS = GetNoteGlobalTarget("dyn")
'    ActiveSheet.Range("EQ4").Value = IIf(RS = -555, "", RS)
    
    Application.Calculation = xlCalculationAutomatic
    On Error Resume Next
    ThisWorkbook.sheets("RATING").Shapes("ZoneTexte 37").TextFrame2.TextRange.Characters.text = (val(ThisWorkbook.sheets("Calculs").Range("seuilvA").Value) * 100) & "%"
    ThisWorkbook.sheets("RATING").Shapes("ZoneTexte 39").TextFrame2.TextRange.Characters.text = (val(ThisWorkbook.sheets("Calculs").Range("seuilrA").Value) * 100) & "%"
    On Error GoTo 0
    Call getTotalColor
End Sub


