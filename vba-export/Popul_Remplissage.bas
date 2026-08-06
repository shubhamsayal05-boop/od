Attribute VB_Name = "Popul_Remplissage"
' Cellule Office
Option Explicit
Option Compare Text
Function Remplissage_Population(onglet As String)
    Dim r As Range
    Dim NEVENTS As Variant
    Dim Multishift As Integer
    Dim SDV As String
    Dim trouve As Range
    Dim k As Long
    Dim tabV() As String
    Dim nn As Long
    Dim st As String
    Dim Ncrit
    Dim critere()
    Dim lastRow As Integer
    Dim rowCopy As Integer, o As Long, n As Long
    Dim i As Integer, j As Integer
    Dim GearOld As Integer, GearNew As Integer, vT As Variant
    Dim rangeToCopy As Range, PedalF As String, v As Variant, transfertPedal() As Variant, colPedal As Integer
    
    With ThisWorkbook.sheets(onglet)
   
        Call showC3(onglet, "driv")
        NEVENTS = ThisWorkbook.sheets("Structure").Range("N1")
        SDV = toSQL(onglet)
       
        Ncrit = SDV2Ncrit(onglet)
        ReDim critere(Ncrit + 1)

        Multishift = 0
        On Error Resume Next

        Call FiltresOff
        PedalF = ""
        ThisWorkbook.sheets("DATA").Rows(1).AutoFilter Field:=3, Criteria1:=onglet
        lastRow = ThisWorkbook.sheets("DATA").Range("A65000").End(xlUp).row
        If lastRow = 1 Then
            Exit Function
        End If

        For rowCopy = 2 To NEVENTS
            GearOld = val(ThisWorkbook.sheets("DATA").Cells(rowCopy, 5).Value)
            GearNew = val(ThisWorkbook.sheets("DATA").Cells(rowCopy, 4).Value)
            If Multishift > 0 Then
                If GearOld <> 0 And GearNew <> 0 And Abs(GearOld - GearNew) <> Multishift Then
                    ThisWorkbook.sheets("DATA").Cells(rowCopy, 5).EntireRow.Hidden = True
                End If
            End If
        Next rowCopy
        
       
        st = getNumberRow(onglet)
        If st = "" Then Exit Function
        o = val(Split(st, ";")(1))
        n = val(Split(st, ";")(0)) + 1
        Do While o >= n
            Set r = ThisWorkbook.sheets("structure").Cells(n, 2)
            If StrComp(ThisWorkbook.sheets("structure").Cells(r.row, 3).Value, "criteria", vbTextCompare) = 0 Or StrComp(ThisWorkbook.sheets("structure").Cells(r.row, 3).Value, "data", vbTextCompare) = 0 Then
                If ThisWorkbook.sheets("structure").Cells(r.row, 3).Value = "criteria" Then
                     
                     critere(k) = ThisWorkbook.sheets("structure").Cells(r.row, 4).Value
                     If Len(ThisWorkbook.sheets("structure").Cells(r.row, 5).Value) > 0 Then
                        tabV = Split(replace(ThisWorkbook.sheets("structure").Cells(r.row, 5).Value, ", ", ","), ",")
                        Set rangeToCopy = ThisWorkbook.sheets("DATA").Range(nomcol(tabV(0), tabV(1)) & "2:" & nomcol(tabV(0), tabV(1)) & NEVENTS)
        '                rangeToCopy.SpecialCells(xlCellTypeVisible).Copy
                     Else
                        If Not ThisWorkbook.sheets("DATA").Range("A1:BIZ1").Find(critere(k), , , xlWhole) Is Nothing Then
                           Set rangeToCopy = ThisWorkbook.sheets("DATA").Range("A1:BIZ1").Find(critere(k), , , xlWhole)
                           Set rangeToCopy = ThisWorkbook.sheets("DATA").Range(rangeToCopy.Offset(1, 0), rangeToCopy.Offset(lastRow - 1, 0))
        '                   ThisWorkbook.Sheets("DATA").Range(SELECTION, SELECTION.offset(lastRow - 1, 0)).SpecialCells(xlCellTypeVisible).Copy
                        End If
                    End If
                    
                    Set trouve = ThisWorkbook.sheets(onglet).Rows(6).Cells.Find(What:=critere(k), LookAt:=xlWhole)
                    If Not trouve Is Nothing Then
                        Set trouve = Nothing
                        vT = CopyColumn(rangeToCopy)
                        colPedal = ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:=critere(k), LookAt:=xlWhole).Column
                       
                        ThisWorkbook.sheets(onglet).Range(ThisWorkbook.sheets(onglet).Cells(7, colPedal), ThisWorkbook.sheets(onglet).Cells(UBound(vT) + 7, colPedal)) = Application.Transpose(vT)

                    End If
                     k = k + 1
                     
                Else
                        tabV = Split(replace(ThisWorkbook.sheets("structure").Cells(r.row, 5).Value, ", ", ","), ",")
                        Set trouve = ThisWorkbook.sheets(onglet).Rows(6).Cells.Find(What:=ThisWorkbook.sheets("structure").Cells(r.row, 4).Value, LookAt:=xlWhole)
                        If Not trouve Is Nothing Then
                            Set trouve = Nothing
                            colPedal = ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:=ThisWorkbook.sheets("structure").Cells(r.row, 4).Value, LookAt:=xlWhole).Column
                            vT = CopyColumn(ThisWorkbook.sheets("DATA").Range(nomcol(tabV(0), tabV(1)) & "2:" & nomcol(tabV(0), tabV(1)) & NEVENTS))
                            ThisWorkbook.sheets(onglet).Range(ThisWorkbook.sheets(onglet).Cells(7, colPedal), ThisWorkbook.sheets(onglet).Cells(UBound(vT) + 7, colPedal)) = Application.Transpose(vT)
                            
                          
                        End If
                         If InStr(1, UCase(ThisWorkbook.sheets("structure").Cells(r.row, 4).Value), "THROTTLE POSITION") <> 0 Or InStr(1, UCase(replace(ThisWorkbook.sheets("structure").Cells(r.row, 4).Value, "é", "E")), "POSITION PEDALE") Then
                                PedalF = ThisWorkbook.sheets("structure").Cells(r.row, 4).Value
                        End If
                End If
                
        
            End If
        
       
            n = n + 1
        Loop
       
        

       
       ThisWorkbook.sheets("DATA").Range("A2:A" & NEVENTS).SpecialCells(xlCellTypeVisible).Copy Destination:=.Cells(7, ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:="Id BdD", LookAt:=xlWhole).Column)
       
       
        Application.CutCopyMode = False
         
         If getLever(onglet) = True Then
                For i = 1 To UBound(LeverCount)
                     j = 2
                     While Len(ThisWorkbook.Worksheets("CONFIGURATIONS").Range("A" & j)) > 0
                         .Columns(ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:=LeverCount(i), LookAt:=xlWhole).Column).replace _
                         What:=ThisWorkbook.Worksheets("CONFIGURATIONS").Range("A" & j), _
                         Replacement:=ThisWorkbook.Worksheets("CONFIGURATIONS").Range("B" & j), LookAt:=xlWhole, _
                         SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                         ReplaceFormat:=False
                         
                         
                         
                         j = j + 1
                     Wend
                Next i
         End If
         
 
    
    End With
End Function

Function CopyColumn(plageColumn As Range) As Variant
    Dim cols As String
    Dim cS(0)
    Dim r As Range
    
    cols = "||[@]"
    For Each r In plageColumn.SpecialCells(xlCellTypeVisible)
            If r.EntireColumn.Hidden = False Then
                If cols = "||[@]" Then cols = fixTheDecimalSeparator(r.Value) Else cols = cols & "#" & fixTheDecimalSeparator(r.Value)
            End If
    Next r
    If InStr(1, cols, "#") = 0 Then
        cS(0) = cols
        CopyColumn = cS
    Else
        CopyColumn = Split(cols, "#")
    End If
End Function












