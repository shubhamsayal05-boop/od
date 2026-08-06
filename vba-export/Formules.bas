Attribute VB_Name = "Formules"
Option Explicit
Sub DelVEHICLE()
deleteVeh.Show vbModeless

End Sub

Function remplirFormule()
    Dim r As Range
    Dim D As Long
    Dim f As Long
    
    ResetValues
    D = StartConFig()
    If D = 0 Then Exit Function
    f = LastRS(D)
    If f = 0 Then Exit Function
    With ThisWorkbook.Worksheets("Calculs")
        Set r = .Range("B5")
        While Len(r.Value) > 0
            If Len(r.Offset(1, 0)) > 0 Then
                r.Offset(0, 1).Value = findV(r.Value, D, f, 2)
                r.Offset(0, 2).Value = findV(r.Value, D, f, 3)
                r.Offset(0, 3).Value = findV(r.Value, D, f, 4)
                r.Offset(0, 4).Value = findV(r.Value, D, f, 5)
            End If
            Set r = r.Offset(1, 0)
        Wend
    End With
End Function
Function StartConFig() As Long
    Dim lastr As Long
    Dim r As Range
    Dim Engine As String, Gearbox As String, NbGear As String, Area As String
    Dim OK(4) As Boolean
    Dim i As Integer
    StartConFig = 0
    With ThisWorkbook.sheets("HOME")
            Engine = .Range("Fuel")
            If InStr(1, .Range("Gears"), " ") <> 0 And .Range("Gears") <> "MANUAL GEARBOX" Then
                Gearbox = Left(.Range("Gears"), (InStr(1, .Range("Gears"), " ") - 1))
            Else
                Gearbox = .Range("Gears")
            End If
            NbGear = .Range("H23")
            Area = .Range("Area")
    End With
  
    OK(1) = False
    OK(2) = False
    OK(3) = False
    OK(4) = False
    With ThisWorkbook.sheets("POWERTRAIN")
            Set r = .Range("A3")
            lastr = .Range("B65000").End(xlUp).row
          
            While Application.CountA(.Rows(r.row).EntireRow) > 0 And r.row <= lastr
                If UCase(.Range("A" & r.row)) = "TITRE CONFIG" Then
                    For i = 2 To .Cells(r.Offset(1, 1).row, .Columns.Count).End(xlToLeft).Column
                        If UCase(Engine) = UCase(r.Offset(1, i - 1)) Then
                            If UCase(r.Offset(2, i - 1)) = "X" Then OK(1) = True
                        End If
                    Next i
                  
                    For i = 2 To .Cells(r.Offset(3, 1).row, .Columns.Count).End(xlToLeft).Column
                        If UCase(Gearbox) = UCase(r.Offset(3, i - 1)) Then
                            If UCase(r.Offset(4, i - 1)) = "X" Then OK(2) = True
                        End If
                    Next i
               
                    For i = 2 To .Cells(r.Offset(5, 1).row, .Columns.Count).End(xlToLeft).Column
                        If UCase(NbGear) = UCase(r.Offset(5, i - 1)) Then
                            If UCase(r.Offset(6, i - 1)) = "X" Then OK(3) = True
                        End If
                        
                    Next i
                  
                    For i = 2 To .Cells(r.Offset(7, 1).row, .Columns.Count).End(xlToLeft).Column
                        If UCase(Area) = UCase(r.Offset(7, i - 1)) Then
                            If UCase(r.Offset(8, i - 1)) = "X" Then OK(4) = True
                        End If
                        
                    Next i
                  
                    If OK(1) = True And OK(2) = True And OK(3) = True And OK(4) = True Then
                        StartConFig = r.Offset(9, 0).row
                        Exit Function
                    Else
                        OK(1) = False
                        OK(2) = False
                        OK(3) = False
                        OK(4) = False
                    End If
                    
                End If
                Set r = r.Offset(1, 0)
            Wend
    End With
End Function
Function findV(recherche As String, D As Long, f As Long, c As Integer)
    On Error Resume Next
    findV = Application.WorksheetFunction.VLookup(recherche, sheets("POWERTRAIN").Range("A" & D & ":I" & f), c, False)
   If ERR.Number <> 0 Then
        ERR.Clear
        findV = 0
   End If
End Function
Function LastRS(i As Long) As Long
    Dim r As Range
   LastRS = 0
     With ThisWorkbook.Worksheets("PowerTrain")
            Set r = .Range("A" & i)
            While r.row <= .Range("A65000").End(xlUp).row
                If r.Value = "Titre config" Then
                       LastRS = r.row - 2
                       Exit Function
                ElseIf r.row = .Range("A65000").End(xlUp).row Then
                      LastRS = r.row - 1
                End If
                Set r = r.Offset(1, 0)
            Wend
            
        End With
End Function

Function ResetValues()
Dim r As Range
With ThisWorkbook.Worksheets("Calculs")
        Set r = .Range("B5")
        While Len(r.Value) > 0
            Set r = r.Offset(1, 0)
        Wend
        .Range("C5:F" & (r.row - 2)).Value = 0
End With
End Function
Sub AddColMode()
Dim c As Range
With ThisWorkbook.Worksheets("CONFIGURATIONS")
        Set c = .Range("COLMODESCONFIG")
        Set c = c.Offset(1, 0)
        While c.Value <> ""
            Set c = c.Offset(1, 0)
        Wend
        If .Range("A" & c.row & ":B" & c.row).MergeCells = False Then
            .Rows(c.row + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            .Range("A" & c.row & ":B" & c.row).Borders(1).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(2).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(3).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(4).LineStyle = xlContinuous
           
        End If
   End With
   
End Sub
Sub AddDMU()
Dim c As Range
With ThisWorkbook.Worksheets("CONFIGURATIONS")
        Set c = .Range("DMU")
        Set c = c.Offset(1, 0)
        While c.Value <> ""
            Set c = c.Offset(1, 0)
        Wend
        If .Range("A" & c.row & ":B" & c.row).Borders(1).LineStyle = xlNone Then
            .Rows(c.row + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            .Range("A" & c.row & ":B" & c.row).Borders(1).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(2).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(3).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(4).LineStyle = xlContinuous
        End If
   End With
   
End Sub

Sub newRatingSDV()
   ediitSDVName.Show
End Sub

Sub AddsPowertrain()
    AddPowertrain.Show
End Sub

Sub configPowertrain()
    EditPowertrain.Show
End Sub

Sub delsPowertrain()
    delPowertrain.Show
End Sub

Sub AddSeetings()
    AddSetting.Show
End Sub
Sub delSeetings()
    delSetting.Show
End Sub

Sub confSeetings()
  
   On Error Resume Next
    If ActiveCell.Column = 1 And ActiveCell.row > 1 And Len(ActiveCell.Value) > 0 Then EditSeeting.ComboBox2 = ActiveCell.Value
    On Error GoTo 0
    EditSeeting.Show
  
End Sub

Sub AddVERSION()
Dim c As Range
With ThisWorkbook.Worksheets("CONFIGURATIONS")
        Set c = .Range("VERSION")
        Set c = c.Offset(1, 0)
        While c.Value <> ""
            Set c = c.Offset(1, 0)
        Wend
        If .Range("A" & c.row & ":B" & c.row).MergeCells = False Then
            .Rows(c.row + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            .Range("A" & c.row & ":B" & c.row).Borders(1).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(2).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(3).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(4).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Merge
        End If
   End With
   
End Sub


Sub AddVEHICLE()
'Dim c As Range
'With ThisWorkbook.Worksheets("CONFIGURATIONS")
'        Set c = .Range("VEHICLE")
'        Set c = c.Offset(1, 0)
'        While c.Value <> ""
'            Set c = c.Offset(1, 0)
'        Wend
'        If .Range("A" & c.Row & ":B" & c.Row).MergeCells = False Then
'            .Rows(c.Row + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
'            .Range("A" & c.Row & ":B" & c.Row).Borders(1).LineStyle = xlContinuous
'            .Range("A" & c.Row & ":B" & c.Row).Borders(2).LineStyle = xlContinuous
'            .Range("A" & c.Row & ":B" & c.Row).Borders(3).LineStyle = xlContinuous
'            .Range("A" & c.Row & ":B" & c.Row).Borders(4).LineStyle = xlContinuous
'            .Range("A" & c.Row & ":B" & c.Row).Merge
'        End If
'   End With
   defineVeh.Show vbModeless
End Sub
Sub AddAREA()
Dim c As Range
With ThisWorkbook.Worksheets("CONFIGURATIONS")
        Set c = .Range("AREA")
        Set c = c.Offset(1, 0)
        While c.Value <> ""
            Set c = c.Offset(1, 0)
        Wend
        If (c.row - (.Range("AREA").row + 1)) >= totalLimits(3) Then
            MsgBox "Attention Limite", vbCritical, "ODRIV"
            Exit Sub
        End If
        If .Range("A" & c.row & ":B" & c.row).MergeCells = False Then
            .Rows(c.row + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            .Range("A" & c.row & ":B" & c.row).Borders(1).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(2).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(3).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(4).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Merge
          
        End If
   End With
   
End Sub
Sub AddENGINE()
Dim c As Range
With ThisWorkbook.Worksheets("CONFIGURATIONS")
        Set c = .Range("ENGINE")
        Set c = c.Offset(1, 0)
        While c.Value <> ""
            Set c = c.Offset(1, 0)
        Wend
         If (c.row - (.Range("ENGINE").row + 1)) >= totalLimits(1) Then
            MsgBox "Attention Limite", vbCritical, "ODRIV"
            Exit Sub
        End If
        If .Range("A" & c.row & ":B" & c.row).MergeCells = False Then
            .Rows(c.row + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            .Range("A" & c.row & ":B" & c.row).Borders(1).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(2).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(3).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(4).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Merge
        End If
   End With
   
End Sub
Sub AddGEARBOX()
Dim c As Range
With ThisWorkbook.Worksheets("CONFIGURATIONS")
        Set c = .Range("GEARBOX")
        Set c = c.Offset(1, 0)
        While c.Value <> ""
            Set c = c.Offset(1, 0)
        Wend
         If (c.row - (.Range("GEARBOX").row + 1)) >= totalLimits(2) Then
            MsgBox "Attention Limite", vbCritical, "ODRIV"
            Exit Sub
        End If
        If .Range("A" & c.row & ":F" & c.row).Borders(1).LineStyle = xlNone Then
            .Rows(c.row + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            .Range("A" & c.row & ":F" & c.row).Borders(1).LineStyle = xlContinuous
            .Range("A" & c.row & ":F" & c.row).Borders(2).LineStyle = xlContinuous
            .Range("A" & c.row & ":F" & c.row).Borders(3).LineStyle = xlContinuous
            .Range("A" & c.row & ":F" & c.row).Borders(4).LineStyle = xlContinuous
        End If
   End With

End Sub
Sub AddNBGEAR()
Dim c As Range
With ThisWorkbook.Worksheets("CONFIGURATIONS")
        Set c = .Range("NBGEAR")
        Set c = c.Offset(1, 0)
        While c.Value <> ""
            Set c = c.Offset(1, 0)
        Wend
        If (c.row - (.Range("NBGEAR").row + 1)) >= totalLimits(4) Then
            MsgBox "Attention Limite", vbCritical, "ODRIV"
            Exit Sub
        End If
        If .Range("A" & c.row & ":B" & c.row).MergeCells = False Then
            .Rows(c.row + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            .Range("A" & c.row & ":B" & c.row).Borders(1).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(2).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(3).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(4).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Merge
        End If
   End With
   
End Sub
Sub AddNBMode()
Dim c As Range
With ThisWorkbook.Worksheets("CONFIGURATIONS")
        Set c = .Range("MODESCONFIG")
        Set c = c.Offset(1, 0)
        While c.Value <> ""
            Set c = c.Offset(1, 0)
        Wend
        If .Range("A" & c.row & ":C" & c.row).MergeCells = False Then
            .Rows(c.row + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            .Range("A" & c.row & ":C" & c.row).Borders(1).LineStyle = xlContinuous
            .Range("A" & c.row & ":C" & c.row).Borders(2).LineStyle = xlContinuous
            .Range("A" & c.row & ":C" & c.row).Borders(3).LineStyle = xlContinuous
            .Range("A" & c.row & ":C" & c.row).Borders(4).LineStyle = xlContinuous
           
        End If
   End With
   
End Sub
Sub adMILESTONE()
Dim c As Range
With ThisWorkbook.Worksheets("CONFIGURATIONS")
        Set c = .Range("MILESTONE")
        Set c = c.Offset(1, 0)
        While c.Value <> ""
            Set c = c.Offset(1, 0)
        Wend
        If .Range("A" & c.row & ":B" & c.row).Borders(1).LineStyle = xlNone Then
            .Rows(c.row + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            .Range("A" & c.row & ":B" & c.row).Borders(1).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(2).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(3).LineStyle = xlContinuous
            .Range("A" & c.row & ":B" & c.row).Borders(4).LineStyle = xlContinuous
        End If
   End With
   
End Sub

Function totalLimits(id As Integer)
    If id = 1 Then
        totalLimits = 11
   ElseIf id = 2 Then
        totalLimits = 11
  ElseIf id = 3 Then
        totalLimits = 12
  ElseIf id = 4 Then
        totalLimits = 12
   End If
End Function




