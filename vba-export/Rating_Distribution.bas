Attribute VB_Name = "Rating_Distribution"
Option Explicit

Sub Distributions(ByVal onglet As String, Prt As String)
    
    Dim NEVENTS As Integer
    Dim derniereLigne As Long
    
    derniereLigne = TotEventSheet(onglet)
    If Prt = "driv" Then ThisWorkbook.sheets(onglet).Range("G8") = derniereLigne - 6
    If Prt = "dyn" Then ThisWorkbook.sheets(onglet).Range("BN8") = derniereLigne - 6
    NEVENTS = ThisWorkbook.sheets(onglet).Range("G8").Value

    With ThisWorkbook.sheets(onglet)
        If Prt = "driv" Then
        .Range("G11") = Application.WorksheetFunction.CountIf(.Range("Q7:O" & NEVENTS + 6), "GREEN")
        .Range("G14") = Application.WorksheetFunction.CountIf(.Range("Q7:O" & NEVENTS + 6), "YELLOW")
        .Range("G17") = Application.WorksheetFunction.CountIf(.Range("Q7:O" & NEVENTS + 6), "RED")
        .ChartObjects("RECAPDRI").Chart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)
        .ChartObjects("RECAPDRI").Chart.SeriesCollection(2).Points(1).Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
        .ChartObjects("RECAPDRI").Chart.SeriesCollection(3).Points(1).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
       ElseIf Prt = "dyn" Then
        .Range("BN11") = Application.WorksheetFunction.CountIf(.Range("BX7:BV" & NEVENTS + 6), "GREEN")
        .Range("BN14") = Application.WorksheetFunction.CountIf(.Range("BX7:BV" & NEVENTS + 6), "YELLOW")
        .Range("BN17") = Application.WorksheetFunction.CountIf(.Range("BX7:BV" & NEVENTS + 6), "RED")
        .ChartObjects("RECAPDYN").Chart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)
        .ChartObjects("RECAPDYN").Chart.SeriesCollection(2).Points(1).Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
        .ChartObjects("RECAPDYN").Chart.SeriesCollection(3).Points(1).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
       End If
    End With
End Sub








