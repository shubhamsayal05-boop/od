Attribute VB_Name = "SetAsTarget"
Option Explicit
Private dict As New Dictionary
Private v As Variant
Private key As Variant

Function RemplirDictionnaire()
    Dim plage As Range
    Dim cell As Range
    Dim i As Integer
    Dim j As Long
    Dim l As Long
'    l = getLastRowRating
   ' v = ""
    Set dict = CreateObject("Scripting.Dictionary")
    dict.RemoveAll
   
    With ThisWorkbook.sheets("TARGET VEHICLE")
        l = .Cells(Rows.Count, 1).End(xlUp).row
        Set plage = .Range("A2:A" & l)
        For Each cell In plage
            If Not isEmpty(cell.Value) Then
                v = cell.Value & "," & cell.Offset(0, 1).Value & "," & cell.Offset(0, 2).Value & "," & cell.Offset(0, 3).Value
                If Not dict.Exists(v) Then
                    dict.Add v, cell.row
                End If
            End If
        Next cell
    End With
End Function
Sub chargement()
    Dim l As Long
    Dim j As Long
    Dim val1, val2, val3
    Dim r As Range
    Dim col As String
    Dim IR
    l = getLastRowRating
    With ThisWorkbook.Worksheets("RATING")
        If Not .Rows("21:22").Find(What:="Responsiveness Index", LookAt:=xlWhole) Is Nothing Then

                Set r = .Rows("21:22").Find(What:="Responsiveness Index", LookAt:=xlWhole)
                col = Split(Cells(22, r.Column).Address, "$")(1)
        End If
      
        Call supprimerdoublons
        Call RemplirDictionnaire
        For j = 23 To l
             If .Rows(j).Hidden = False And .Cells(j, 4) <> "" Then
                    val1 = .Cells(j, 4)
                    val2 = .Range("M" & j)
                    val3 = .Range(col & j)
                    Call verif(val1, val2, val3)
             End If
         Next j
         val1 = "Rate of low points"
         IR = .Rows("10:10").Find(What:="Tested vehicle", LookAt:=xlWhole).Column
         val2 = .Cells(12, IR)
         IR = .Rows("16:16").Find(What:="Tested vehicle", LookAt:=xlWhole).Column
         val3 = .Cells(18, IR)
         Call verif(val1, val2, val3)
     End With
     Call Remplir_Configurations
     With ThisWorkbook.Worksheets("RATING")
        If sheets("HOME").Range("Milestone") <> 4 Then
            .Columns("F:F").Hidden = False
            .Columns(get_col_name("Driveability Index", -3)).Hidden = False
            .Columns(get_col_name("Drivability Lowest Events", 4)).Hidden = False
        Else
            .Columns("F:F").Hidden = True
            .Columns(get_col_name("Driveability Index", -3)).Hidden = True
            .Columns(get_col_name("Drivability Lowest Events", 4)).Hidden = True
        End If
    End With
End Sub
Function verif(val1, val2, val3)
   v = val1 & "," & ThisWorkbook.sheets("HOME").Range("DriveVersion") & "," & ThisWorkbook.sheets("HOME").Range("Project") & "," & ThisWorkbook.sheets("HOME").Range("Mode")
   With ThisWorkbook.sheets("TARGET VEHICLE")
        If dict.Exists(v) Then
            .Cells(dict(v), 5) = val2
            .Cells(dict(v), 6) = val3
       Else
        Call insert_row(val1, val2, val3)
'        Call RemplirDictionnaire
        dict.Add v, ThisWorkbook.sheets("TARGET VEHICLE").Cells(Rows.Count, 1).End(xlUp).row
       End If
   End With
End Function

Function supprimerdoublons()
Dim l As Long
 With ThisWorkbook.sheets("TARGET VEHICLE")
         l = .Cells(Rows.Count, 1).End(xlUp).row
        .Range("A1:F" & l).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6), Header:=xlYes
        
        .Range("A1:F" & l).Borders(1).LineStyle = xlContinuous
        .Range("A1:F" & l).Borders(2).LineStyle = xlContinuous
        .Range("A1:F" & l).Borders(3).LineStyle = xlContinuous
        .Range("A1:F" & l).Borders(4).LineStyle = xlContinuous
 End With
End Function


Function insert_row(val1, val2, val3)
    Dim l, l1 As Long
    With ThisWorkbook.sheets("TARGET VEHICLE")
         l = .Cells(Rows.Count, 1).End(xlUp).row
         l1 = l + 1
'         .Range("A" & l & ":F" & l).Copy
'          Application.DisplayAlerts = False
'          .Range("A" & l1).EntireRow.Insert
'          .Range("A" & l1 & ":F" & l1).ClearContents
          
          .Range("A" & l1) = val1
          .Range("B" & l1) = ThisWorkbook.sheets("HOME").Range("DriveVersion")
          .Range("C" & l1) = ThisWorkbook.sheets("HOME").Range("Project")
          .Range("D" & l1) = ThisWorkbook.sheets("HOME").Range("Mode")
          .Range("E" & l1) = val2
          .Range("F" & l1) = val3
    End With
End Function


Function Remplir_Configurations()
    Dim c As Range
    Dim T_veh As Boolean
    
    T_veh = False
    With ThisWorkbook.Worksheets("CONFIGURATIONS")
        Set c = .Range("VEHICLE")
        Set c = c.Offset(1, 0)
        While c.Value <> ""
            If c.Value = ThisWorkbook.sheets("HOME").Range("Project").Value Then
                T_veh = True
            End If
            Set c = c.Offset(1, 0)
        Wend
        If .Range("A" & c.row & ":B" & c.row).MergeCells = False Then
            If T_veh = False Then '.Range("A" & c.row - 1).Value <> ThisWorkbook.sheets("HOME").Range("Project") Then
                sheets("CONFIGURATIONS").Visible = xlSheetVisible
                 sheets("CONFIGURATIONS").Activate
                Call hideShowTarget(False)
                Call defineVeh.addVehRating2(ThisWorkbook.sheets("HOME").Range("Project"))
                Call hideShowTarget(True)
                Call defineVeh.addVehGraphStatus2(ThisWorkbook.sheets("HOME").Range("Project"))
                .Rows(c.row + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                .Range("A" & c.row & ":B" & c.row).Borders(1).LineStyle = xlContinuous
                .Range("A" & c.row & ":B" & c.row).Borders(2).LineStyle = xlContinuous
                .Range("A" & c.row & ":B" & c.row).Borders(3).LineStyle = xlContinuous
                .Range("A" & c.row & ":B" & c.row).Borders(4).LineStyle = xlContinuous
                .Range("A" & c.row & ":B" & c.row).Merge
                .Range("A" & c.row & ":B" & c.row) = ThisWorkbook.sheets("HOME").Range("Project")
                Call UpdateYVal2
                Call defineVeh.addVehTotPoint
                sheets("CONFIGURATIONS").Visible = xlSheetVeryHidden

            End If
        End If
        sheets("RATING").Activate
        sheets("RATING").Shapes("UpdateTargetButton").Visible = False
    End With
End Function


Private Sub UpdateYVal2()
 
    Dim ws As Worksheet

    Dim lastRow As Long

    Dim i As Long, j As Long

    Dim GI As Collection

    Dim IR As Collection

    Dim Values As Collection

    Dim n As Long

    Dim stepVal As Double

    Dim minVal As Double, maxVal As Double

    minVal = 4

    maxVal = 99
 
    Set ws = ThisWorkbook.sheets("Graph_status")

    Set GI = New Collection

    Set IR = New Collection

    Set Values = New Collection
 
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
 
    

    For i = 2 To lastRow

        If ws.Cells(i, 1).Value = "Global index" Then

            GI.Add i + 1

        ElseIf ws.Cells(i, 1).Value = "index rouge" Then

            IR.Add i - 2

        End If

    Next i
 
   

    For i = 1 To GI.Count

        n = IR(i) - GI(i) + 1

        stepVal = (maxVal - minVal) / (n - 1)

        If i = 1 Then

            

            For j = 0 To n - 1

                ws.Cells(GI(i) + j, 4).Value = minVal + stepVal * j

                Values.Add minVal + stepVal * j

            Next j
 
        Else

            

            For j = 1 To n

                ws.Cells(GI(i) + j - 1, 4).Value = Values(j)

            Next j

        End If

    Next i
 
End Sub

 






