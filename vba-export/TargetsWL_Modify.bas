Attribute VB_Name = "TargetsWL_Modify"
' Cellule Office

Option Explicit

Sub MajAll_WTP(ByVal Mode As String, ByVal nom As String)
    Dim v
    Dim i As Long
    ThisWorkbook.Worksheets("HOME").Range("Moniteur").Interior.color = RGB(255, 0, 0)
    
    v = ThisWorkbook.sheets("structure").UsedRange.Columns(2).Value
    For i = 2 To UBound(v, 1)
        If Len(v(i, 1)) > 0 And sheetExists(v(i, 1)) = True Then
           Call affect_WTP(v(i, 1))
        End If
    Next i
    Erase v

    Call Moniteur("Targets dataset """ & nom & " - Mode " & Mode & """ has been applied to your project.")
End Sub

Sub affect_WTP(ByVal onglet As String, Optional Prt As String)
    Dim r As Range
    Dim v
    Dim i As Long
    Dim founded As Boolean
    Dim targetRange As String
    Dim Mode As String
    Dim Fuel As String
    Dim version As String
    Dim priority As String
    Dim criteria As String
    Dim rt(1) As String
    Dim j As Integer
    Dim Target As Single, waterline As Single, Criticity As Integer
    Dim rang As String
    Dim iDPart As Integer
    
    If Len(Prt) = 0 Then
            rt(0) = "A6:BA6"
            rt(1) = "BT6:GG6"
            Call showC3(onglet)
        
            v = ThisWorkbook.sheets("TARGETS").UsedRange.Value
            targetRange = "PREMIUM"
            Mode = ThisWorkbook.sheets("HOME").Range("Mode").Value
            version = ThisWorkbook.sheets("HOME").Range("DriveVersion").Value
            If version <> "V3.8" Then
                Fuel = ""
            Else
                Fuel = ThisWorkbook.sheets("HOME").Range("Fuel").Value
            End If
            priority = ThisWorkbook.sheets("HOME").Range("Prestation").Value
            Call deleteAllB(onglet, 0)
            Call deleteAllB(onglet, 1)
            For i = 1 To UBound(v, 1)
                If StrComp(v(i, 1), onglet, vbTextCompare) = 0 And StrComp(v(i, 3), targetRange, vbTextCompare) = 0 And InStr(1, UCase(";" & v(i, 4) & ";"), ";" & Mode & ";") <> 0 And StrComp(v(i, 5), Fuel, vbTextCompare) = 0 And StrComp(v(i, 6), version, vbTextCompare) = 0 Then
                    founded = True
        
                    waterline = v(i, 7)
                    Target = v(i, 8)
        
        '            If StrComp(priority, "DYNAMIC", vbTextCompare) = 0 Then
        '                Criticity = V(i, 9)
        '            ElseIf StrComp(priority, "DRIVABILITY", vbTextCompare) = 0 Then
        '                Criticity = V(i, 10)
        '            End If
        '
        
                    criteria = v(i, 2)
                    ' trouver cellule criteria sur onglet, maj valeurs. ici
                    
                    For j = 0 To 1
                         Set r = ThisWorkbook.sheets(onglet).Range(rt(j)).Cells.Find(What:=criteria, LookAt:=xlWhole)
                         
                         If Not r Is Nothing Then
                             If j = 1 Then
                                r.Offset(-1, 0).Value = IIf(Len(Target) = 0 And Len(waterline) = 0, 3, v(i, 9))
                             Else
                                r.Offset(-1, 0).Value = IIf(Len(Target) = 0 And Len(waterline) = 0, 3, v(i, 10))
                             End If
                             r.Offset(-2, 0).Value = Target
                             r.Offset(-3, 0).Value = waterline
                         
                             Set r = Nothing
                        End If
                   Next j
                End If
            Next i
            
          
        
            Erase v
        
        '    Application.ScreenUpdating = True
     Else
            If Prt = "driv" Then
                rang = "A6:BA6"
                iDPart = 0
            Else
                rang = "BT6:GG6"
                 iDPart = 1
            End If
           
            Call showC3(onglet, Prt)
        
            v = ThisWorkbook.sheets("TARGETS").UsedRange.Value
            targetRange = "PREMIUM"
            Mode = ThisWorkbook.sheets("HOME").Range("Mode").Value
            version = ThisWorkbook.sheets("HOME").Range("DriveVersion").Value
            If version <> "V3.8" Then
                Fuel = ""
            Else
                Fuel = ThisWorkbook.sheets("HOME").Range("Fuel").Value
            End If
            priority = ThisWorkbook.sheets("HOME").Range("Prestation").Value
            
            
            Call deleteAllB(onglet, iDPart)
           
            
            For i = 1 To UBound(v, 1)
                If StrComp(v(i, 1), onglet, vbTextCompare) = 0 And StrComp(v(i, 3), targetRange, vbTextCompare) = 0 And InStr(1, UCase(";" & v(i, 4) & ";"), ";" & Mode & ";") <> 0 And StrComp(v(i, 5), Fuel, vbTextCompare) = 0 And StrComp(v(i, 6), version, vbTextCompare) = 0 Then
                    founded = True
        
                    waterline = v(i, 7)
                    Target = v(i, 8)
                    criteria = v(i, 2)
                    
                    
                         Set r = ThisWorkbook.sheets(onglet).Range(rang).Cells.Find(What:=criteria, LookAt:=xlWhole)
                         
                         If Not r Is Nothing Then
                             If iDPart = 1 Then
                                r.Offset(-1, 0).Value = IIf(Len(Target) = 0 And Len(waterline) = 0, 3, v(i, 9))
                             Else
                                r.Offset(-1, 0).Value = IIf(Len(Target) = 0 And Len(waterline) = 0, 3, v(i, 10))
                             End If
                             r.Offset(-2, 0).Value = Target
                             r.Offset(-3, 0).Value = waterline
                         
                             Set r = Nothing
                        End If
                
                End If
            Next i
            
          
        
            Erase v
        
        '    Application.ScreenUpdating = True
     End If
End Sub

Function deleteAllB(f As String, iDPart As Integer)
        Dim r As Range
        Dim rt(1) As String
        Dim i As Integer
        
        rt(0) = "A3:BA3"
        rt(1) = "BT3:GG3"
        
       
        With ThisWorkbook.sheets(f)
           
                Set r = .Range(rt(iDPart)).Find(What:="Waterline", LookAt:=xlWhole)
                
                If Not r Is Nothing Then
                        Set r = r.Offset(0, 1)
                        While Len(r.Value) > 0 Or Len(r.Offset(1, 0)) > 0 Or Len(r.Offset(2, 0)) > 0
                            r.Value = ""
                            r.Offset(1, 0).Value = ""
                            r.Offset(2, 0).Value = ""
                            Set r = r.Offset(0, 1)
                        Wend
                End If
            
        End With
End Function





