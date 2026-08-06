Attribute VB_Name = "ColorGlobalByStatus"
Function GlobalDrivability()
    Dim r As Range
    Dim v
    Dim i As Long
    Dim iCol As Integer
    Dim sh As Worksheet
    Dim vCol As String
    Dim pr As Integer
    Dim rt As Integer, IR
    
    
      IR = ThisWorkbook.Worksheets("RATING").Rows("10:10").Find(What:="Tested vehicle", LookAt:=xlWhole).Column
     'ThisWorkbook.Sheets("rating").Range("L10").Formula = "=calculs!M39"
    If Not colorGlobalDriv Is Nothing Then
            If colorGlobalDriv.Count <> 0 Then
                    If ThisWorkbook.Worksheets("RATING").Cells(12, IR) < ThisWorkbook.sheets("calculs").Range("seuilvA") Then
                        If (ThisWorkbook.sheets("RATING").Range("RESULTATGLOBAL1") / 100) >= ThisWorkbook.sheets("calculs").Range("seuilrB") Then
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "Low Risk"
                        ElseIf (ThisWorkbook.sheets("RATING").Range("RESULTATGLOBAL1") / 100) < ThisWorkbook.sheets("calculs").Range("seuilrB") Then
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "Medium Risk"
                        Else
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "Low Risk"
                        End If
                    ElseIf ThisWorkbook.Worksheets("RATING").Cells(12, IR) > ThisWorkbook.sheets("calculs").Range("seuilrA") Then
                        If (ThisWorkbook.sheets("RATING").Range("RESULTATGLOBAL1") / 100) >= ThisWorkbook.sheets("calculs").Range("seuilvB") Then
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "High Risk"
                        ElseIf (ThisWorkbook.sheets("RATING").Range("RESULTATGLOBAL1") / 100) < ThisWorkbook.sheets("calculs").Range("seuilrB") Then
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "High Risk"
                        Else
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "High Risk"
                        End If
                    Else
                        If (ThisWorkbook.sheets("RATING").Range("RESULTATGLOBAL1") / 100) >= ThisWorkbook.sheets("calculs").Range("seuilvA") Then
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "Medium Risk"
                        ElseIf (ThisWorkbook.sheets("RATING").Range("RESULTATGLOBAL1") / 100) < ThisWorkbook.sheets("calculs").Range("seuilrB") Then
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "High Risk"
                        Else
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "Medium Risk"
                        End If
                    End If
                    
            End If
   End If

End Function

