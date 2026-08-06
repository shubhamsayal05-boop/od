Attribute VB_Name = "OrderS"
Option Explicit

Function defaultOrder() As Variant
    Dim i As Long, derniereLigne As Long, derniereColonne As Long
    Dim chapOrder(3) As Integer
    Dim v() As String
    Dim saves As Integer
    chapOrder(1) = 0
    chapOrder(2) = 0
    chapOrder(3) = 0
    
  
    On Error GoTo Ers:
    With ThisWorkbook.Worksheets("SDV MANAGER")
            
     
            derniereLigne = 0
            derniereColonne = 7
            For i = 1 To 2
               If .Cells(.Rows.Count, i).End(xlUp).row > derniereLigne Then derniereLigne = .Cells(.Rows.Count, i).End(xlUp).row
            Next i
             
            ReDim v(derniereLigne - 1, 2)
             For i = 2 To derniereLigne
                    If Application.CountA(.Range("A" & i & ":B" & i)) > 0 Then
                             If .Cells(i, 1).Interior.color = 11851260 Then
                                     chapOrder(1) = chapOrder(1) + 1
                                     chapOrder(2) = 0
                                     chapOrder(3) = 0
                                     v(i - 1, 0) = .Range("A" & i).Value
                                     v(i - 1, 1) = intToOrder(chapOrder(1), True)
                                     v(i - 1, 2) = .Range("A" & i).Address
                              Else
                                      chapOrder(2) = chapOrder(2) + 1
                                      saves = chapOrder(3)
                                      chapOrder(3) = 0
                                      v(i - 1, 0) = .Range("A" & i).Value
                                      v(i - 1, 1) = intToOrder(chapOrder(1), True) & "." & intToOrder(chapOrder(2), False)
                                      v(i - 1, 2) = .Range("A" & i).Address
                           End If
                   End If
               Next i
               defaultOrder = v
'               MsgBox UBound(v, 1)
'            .Columns(8).NumberFormat = "@"
'            .Range(.Cells(1, 8), .Cells(UBound(v, 1) + 1, 8)) = Application.Transpose(v)
      
   End With
   
Ers:
    If ERR.Number <> 0 Then
        MsgBox ERR.description, vbCritical, "ODRIV"
        
    End If
End Function
Function intToOrder(i As Integer, order As Boolean) As String
    Dim v(26) As String
    
    v(1) = "a"
    v(2) = "b"
    v(3) = "c"
    v(4) = "d"
    v(5) = "e"
    v(6) = "f"
    v(7) = "g"
    v(8) = "h"
    v(9) = "i"
    v(10) = "j"
    v(11) = "k"
    v(12) = "l"
    v(13) = "m"
    v(14) = "n"
    v(15) = "o"
    v(16) = "p"
    v(17) = "q"
    v(18) = "r"
    v(19) = "s"
    v(20) = "t"
    v(21) = "u"
    v(22) = "v"
    v(23) = "w"
    v(24) = "x"
    v(25) = "y"
    v(26) = "z"
    
   If order = True Then
        intToOrder = UCase(v(i))
   Else
        intToOrder = v(i)
   End If
   
End Function




Function checkCorrect() As String
    Dim v As Variant
    Dim lastRow As Long, i As Long
    Dim getOrder() As String
    Dim colon As Object
    Dim doublon As Object
    Dim chap As String, Schap As String, fonc As String, keys As String
    chap = ""
    Schap = ""
    fonc = ""
    keys = ""

    checkCorrect = ""
    Set colon = CreateObject("Scripting.Dictionary")
    Set doublon = CreateObject("Scripting.Dictionary")
    v = defaultOrder
    For i = 0 To UBound(v, 1)
      If Len(v(i, 1)) > 0 Then
                If InStr(1, v(i, 1), ".") = 0 Then
                    chap = CStr(v(i, 0))
                    
                     If Not colon.Exists(chap) Then
                        colon.Add key:=(chap), Item:=(chap)
                     Else
                         If Not doublon.Exists(chap) Then
                                doublon.Add key:=(chap), Item:=(chap)
                                If checkCorrect = "" Then checkCorrect = "Chapitre : " & chap _
                                Else checkCorrect = checkCorrect & vbCrLf & "Chapitre : " & chap
                         End If
                     End If
                     
                    Schap = ""
                    fonc = ""
                Else
                   
                               fonc = CStr(v(i, 0))
                               keys = fonc
                                If Not colon.Exists(keys) Then
                                   colon.Add key:=(keys), Item:=(keys)
                                Else
                                    If Not doublon.Exists(keys) Then
                                           doublon.Add key:=(keys), Item:=(keys)
                                           If checkCorrect = "" Then checkCorrect = "Fonction : " & fonc _
                                           Else checkCorrect = checkCorrect & vbCrLf & "Fonction : " & fonc
                                    End If
                                End If
                          
                End If
                 
           End If
    Next i

End Function
Function MAJRatingPosition()
     Dim derniereLigne
     Dim i As Long
     Dim j As Long
     Dim p As Long
     Dim o As Integer
     
     Dim colD
     With ThisWorkbook.Worksheets("RATING")
        derniereLigne = 23
        While Len(.Range("D" & derniereLigne)) > 0 Or Len(.Range("B" & derniereLigne)) > 0
           derniereLigne = derniereLigne + 1
        Wend
        derniereLigne = derniereLigne - 1
        
        colD = .Rows("21:22").Find(What:="Responsiveness Lowest Events", LookAt:=xlWhole).Column + 1
        On Error Resume Next
        .Range("B23:" & .Cells(derniereLigne, colD).Address).MergeCells = False
        .Range("B23:" & .Cells(derniereLigne, colD).Address).ClearContents
        .Rows("23:" & derniereLigne).RowHeight = 21.75
        On Error GoTo 0
        
        p = derniereLigne
        j = 23
        While (derniereLigne - 23) < ThisWorkbook.Worksheets("SDV MANAGER").Cells(ThisWorkbook.Worksheets("SDV MANAGER").Rows.Count, 1).End(xlUp).row - 2
                .Rows(p & ":" & p).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                derniereLigne = derniereLigne + 1
        Wend
        
        o = ThisWorkbook.Worksheets("totalPoint").Cells(1, ThisWorkbook.Worksheets("totalPoint").Columns.Count).End(xlToLeft).Column
        While ThisWorkbook.Worksheets("totalPoint").Cells(1, o).Interior.color <> RGB(255, 255, 255)
            o = o + 1
        Wend
                    
        For i = 2 To ThisWorkbook.Worksheets("SDV MANAGER").Cells(ThisWorkbook.Worksheets("SDV MANAGER").Rows.Count, 1).End(xlUp).row
            If InStr(1, ThisWorkbook.Worksheets("SDV MANAGER").Range("B" & i), ".") = 0 Then
                  
                 .Range("B" & j & ":" & .Cells(j, colD).Address).Merge
                 .Range("B" & j & ":" & .Cells(j, colD).Address).MergeCells = False
                 .Range("B" & j & ":" & .Cells(j, colD).Address).Interior.color = RGB(242, 242, 242)
                 .Range("B" & j).Font.Bold = True
                 .Range("B" & j).Font.Size = 16
                 .Range("B" & j).Value = ThisWorkbook.Worksheets("SDV MANAGER").Range("A" & i)
                 .Rows(j & ":" & j).RowHeight = 36
            Else
                ThisWorkbook.Worksheets("totalPoint").Range("S1:" & ThisWorkbook.Worksheets("totalPoint").Cells(1, o).Address).Copy Destination:=.Range("B" & j)
                .Range("D" & j).Value = ThisWorkbook.Worksheets("SDV MANAGER").Range("A" & i)
            End If
            j = j + 1
        Next i
        MsgBox "Terminé", vbInformation, "ODRIV"
    End With
        
End Function






