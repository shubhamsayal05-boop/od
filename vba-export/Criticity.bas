Attribute VB_Name = "Criticity"
Option Explicit

Function F_criticity(ByVal onglet As String, Prt As String) As Variant
    Dim l As Integer
    Dim c As Integer, i As Long
    Dim v As Variant
    Dim tabResultat() As Variant
    Dim j As Integer
    Dim rang As String
    
    If Prt = "driv" Then j = 1
    If Prt = "dyn" Then j = 2
    
    With ThisWorkbook.sheets(onglet)
            If j = 1 Then
                v = .Range(.Cells(7, 14), .Cells(TotEventSheet(onglet), 15)).Value
            Else
                v = .Range(.Cells(7, 73), .Cells(TotEventSheet(onglet), 74)).Value
            End If
            ReDim tabResultat(UBound(v, 1) - 1)
            For i = 1 To UBound(v, 1)
                If Len(v(i, 1)) > 0 And IsNumeric(v(i, 1)) = True Then
                    c = v(i, 1) + 3
                    l = getEquivCriticity(CStr(v(i, 2)))
                    tabResultat(i - 1) = ThisWorkbook.sheets("cfg_criticity").Cells(l, c).Value
                    If LCase(v(i, 2)) = "yellow" Then
                        If checkYellowUpdate(onglet, CStr(v(i, 1)), i + 6, j) <> 0 Then tabResultat(i - 1) = checkYellowUpdate(onglet, CStr(v(i, 1)), i + 6, j)
                    End If
                Else
                     tabResultat(i - 1) = v(i, 1)
                End If
            Next i
            If j = 1 Then
                .Range("M7:M" & UBound(tabResultat) + 7).Value = Application.Transpose(tabResultat)
            Else
                .Range("BT7:BT" & UBound(tabResultat) + 7).Value = Application.Transpose(tabResultat)
            End If
    End With
    
End Function
Function getEquivCriticity(v As String) As Integer
   getEquivCriticity = 0
    If LCase(v) = "green" Then getEquivCriticity = 8
    If LCase(v) = "yellow" Then getEquivCriticity = 6
    If LCase(v) = "red" Then getEquivCriticity = 5
    If LCase(v) = "red +" Then getEquivCriticity = 4
End Function

Function checkYellowUpdate(onglet As String, ind As Integer, lign As Long, optionId)
    Dim colInd As Integer
    checkYellowUpdate = 0
    With ThisWorkbook.sheets(onglet)
        If optionId = 1 Then
            colInd = .Range("A6:BA6").Find("Indice", , , xlPart).Column
        Else
            colInd = .Range("BH6:GG6").Find("Indice", , , xlPart).Column
        End If
        If ind = 1 Then
             If .Cells(lign, colInd) >= ThisWorkbook.sheets("cfg_criticity").Range("D9").Value Then
                checkYellowUpdate = 3
             End If
        ElseIf ind = 2 Then
             If .Cells(lign, colInd) >= ThisWorkbook.sheets("cfg_criticity").Range("E9").Value Then
               checkYellowUpdate = 4
             End If
        End If
    End With
End Function



