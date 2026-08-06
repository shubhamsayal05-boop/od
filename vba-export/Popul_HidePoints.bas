Attribute VB_Name = "Popul_HidePoints"
' Cellule Office

Option Explicit

Sub DisplayRedOnly(ByVal onglet As String)
    Dim lastRow As Integer
    Dim Ncrit As Variant
    Dim k As Integer
    Dim i As Integer

    With ThisWorkbook.sheets(onglet)
        lastRow = .Range("O" & 10000).End(xlUp).row + 1
        Ncrit = SDV2Ncrit(onglet)
        k = 7
        For i = 7 To lastRow
            If .Range("O" & i) = "GREEN" Or ThisWorkbook.sheets(onglet).Range("O" & i) = "YELLOW" Then
                .Range(.Cells(i, 13), .Cells(i, getLastColumnDrivability(onglet))).Clear
            ElseIf .Range("O" & i) = "RED" Or ThisWorkbook.sheets(onglet).Range("O" & i).Value = "RED +" Then
                 .Range(.Cells(i, 13), .Cells(i, getLastColumnDrivability(onglet))).Cut .Range("M" & k)
                k = k + 1
            End If
        Next i

    End With

End Sub

Sub DisplayRedOnlyDyn(ByVal onglet As String)
    Dim lastRow As Integer
    Dim Ncrit As Variant
    Dim k As Integer
    Dim i As Integer

    With ThisWorkbook.sheets(onglet)
        lastRow = .Range("BV" & 10000).End(xlUp).row + 1
        Ncrit = SDV2Ncrit(onglet)
        k = 7
        For i = 7 To lastRow
            If .Range("BV" & i) = "GREEN" Or ThisWorkbook.sheets(onglet).Range("BV" & i) = "YELLOW" Then
                .Range(.Cells(i, 72), .Cells(i, getLastColumnDinamyc(onglet))).Clear
            ElseIf .Range("BV" & i) = "RED" Or ThisWorkbook.sheets(onglet).Range("BV" & i).Value = "RED +" Then
                .Range(.Cells(i, 72), .Cells(i, getLastColumnDinamyc(onglet))).Cut .Range("BT" & k)
                k = k + 1
            End If
        Next i

      
    End With

End Sub


Sub DisplayRedYellow(onglet As String)
    Dim lastRow As Integer
    Dim Ncrit As Variant
    Dim k As Integer
    Dim i As Integer

    With ThisWorkbook.sheets(onglet)
        lastRow = .Range("N" & 10000).End(xlUp).row + 1
        Ncrit = SDV2Ncrit(onglet)

        k = 7
        For i = 7 To lastRow
            If .Range("O" & i).Value = "GREEN" Then
                Range(.Cells(i, 13), .Cells(i, getLastColumnDrivability(onglet))).Clear
            ElseIf .Range("O" & i).Value = "RED" Or ThisWorkbook.sheets(onglet).Range("O" & i).Value = "RED +" Or ThisWorkbook.sheets(onglet).Range("O" & i).Value = "YELLOW" Then
                Range(.Cells(i, 13), .Cells(i, getLastColumnDrivability(onglet))).Cut .Range("M" & k)
                k = k + 1
            End If
        Next
      
    End With

End Sub

Sub DisplayRedYellowDyn(onglet As String)
    Dim lastRow As Integer
    Dim Ncrit As Variant
    Dim k As Integer
    Dim i As Integer

    With ThisWorkbook.sheets(onglet)
        lastRow = .Range("BT" & 10000).End(xlUp).row + 1
        Ncrit = SDV2Ncrit(onglet)

        k = 7
        For i = 7 To lastRow
            If .Range("BV" & i).Value = "GREEN" Then
                Range(.Cells(i, 72), .Cells(i, getLastColumnDinamyc(onglet))).Clear
            ElseIf .Range("BV" & i).Value = "RED" Or ThisWorkbook.sheets(onglet).Range("BV" & i).Value = "RED +" Or ThisWorkbook.sheets(onglet).Range("BV" & i).Value = "YELLOW" Then
                Range(.Cells(i, 72), .Cells(i, getLastColumnDinamyc(onglet))).Cut .Range("BT" & k)
                k = k + 1
            End If
        Next
       
    End With

End Sub





