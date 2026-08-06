Attribute VB_Name = "CusTOMFunction"
Option Explicit

Public Function calculSummCells(Optional VP As Variant) As Double
    Dim i As Long
    Dim s As Variant
    s = VP
    
    i = LasRs
    
   With ThisWorkbook.Worksheets("Calculs")
        calculSummCells = Application.WorksheetFunction.Sum(.Range("C5:C" & i))
   End With
   
End Function
Public Function calculSummPoints(Optional VP As Variant) As Double
    Dim i As Long
    Dim s As Variant
    s = VP
    i = LasRs
   With ThisWorkbook.Worksheets("Calculs")
        calculSummPoints = Application.WorksheetFunction.Sum(.Range("M5:U" & i))
   End With
End Function
Public Function calculSumGamme(Optional VP As Variant) As Double
    Dim i As Long
    Dim s As Variant
    s = VP
    i = LasRs
   With ThisWorkbook.Worksheets("Calculs")
        calculSumGamme = Application.WorksheetFunction.Sum(.Range("Z5:Z" & i))
   End With
End Function
Public Function powerSummCells(r As Range, Optional VP As Variant) As Double
    Dim i As Long
    i = firstLs(r.row)
    Dim s As Variant
    s = VP
   With ThisWorkbook.Worksheets("PowerTrain")
        powerSummCells = Application.WorksheetFunction.Sum(.Range("B" & i & ":B" & (r.row - 1)))
   End With
   
End Function
Public Function ResultTotal(Optional VP As Variant) As Double
   Dim s As Variant
    s = VP
   With ThisWorkbook.Worksheets("Calculs")
        ResultTotal = .Range("allTotRes")
   End With
   
End Function
Function LasRs() As Long
    Dim r As Range
    LasRs = 0
     With ThisWorkbook.Worksheets("Calculs")
            Set r = .Range("B5")
            While Len(r.Value) > 0
                Set r = r.Offset(1, 0)
            Wend
            LasRs = r.row - 2
        End With
End Function

Function firstLs(i As Long) As Long
    Dim r As Range
    firstLs = 0
     With ThisWorkbook.Worksheets("PowerTrain")
            Set r = .Range("A" & i)
            While r.row > 1
                If r.Value = "Operation modes" Then
                        firstLs = r.row + 1
                        Exit Function
                End If
                Set r = r.Offset(-1, 0)
            Wend
            
        End With
End Function



