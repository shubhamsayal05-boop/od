Attribute VB_Name = "getTarget"
Option Explicit

Function getTotalColor()
    Dim v
    Dim i As Long
    Dim valColor As String
    Dim countColor As String
    renit
    v = ThisWorkbook.sheets("structure").UsedRange.Columns(2).Value
    For i = 2 To UBound(v, 1)
        If Len(v(i, 1)) > 0 And sheetExists(v(i, 1)) = True Then
           valColor = getTotColor(ThisWorkbook.Worksheets(v(i, 1)))
           If valColor <> "" Then
                If countColor = "" Then countColor = valColor Else countColor = totalColor(countColor, valColor)
           End If
        End If
    Next i
    Erase v
    If Len(countColor) > 0 Then Call resultatCalcul(countColor)
   
End Function

Function getTotColor(sh As Worksheet) As String
    Dim i As Long
    Dim pr As Integer, rt As Integer
    Dim totV(3) As Integer, totO(3) As Integer, totR(3) As Integer
     
    totV(1) = 0
    totV(2) = 0
    totV(3) = 0
    totO(1) = 0
    totO(2) = 0
    totO(3) = 0
    totR(1) = 0
    totR(2) = 0
    totR(3) = 0
    getTotColor = ""
    If Not sh.Rows(6).Find(What:="Event Priority", LookAt:=xlWhole) Is Nothing _
         And Not sh.Rows(6).Find(What:="Event Rating", LookAt:=xlWhole) Is Nothing Then
               pr = sh.Rows(6).Find(What:="Event Priority", LookAt:=xlWhole).Column
               rt = sh.Rows(6).Find(What:="Event Rating", LookAt:=xlWhole).Column
               For i = 7 To sh.Cells(sh.Rows.Count, rt).End(xlUp).row
                       If sh.Cells(i, rt) = "RED" Or sh.Cells(i, rt) = "RED +" Or sh.Cells(i, rt) = "RED+" Then
                           If CStr(sh.Cells(i, pr)) = "1" Then totR(1) = totR(1) + 1
                           If CStr(sh.Cells(i, pr)) = "2" Then totR(2) = totR(2) + 1
                           If CStr(sh.Cells(i, pr)) = "3" Then totR(3) = totR(3) + 1
                       ElseIf sh.Cells(i, rt) = "GREEN" Then
                           If CStr(sh.Cells(i, pr)) = "1" Then totV(1) = totV(1) + 1
                           If CStr(sh.Cells(i, pr)) = "2" Then totV(2) = totV(2) + 1
                           If CStr(sh.Cells(i, pr)) = "3" Then totV(3) = totV(3) + 1
                       ElseIf sh.Cells(i, rt) Like "YELLOW" Then
                           If CStr(sh.Cells(i, pr)) = "1" Then totO(1) = totO(1) + 1
                           If CStr(sh.Cells(i, pr)) = "2" Then totO(2) = totO(2) + 1
                           If CStr(sh.Cells(i, pr)) = "3" Then totO(3) = totO(3) + 1
                       End If
               Next i
               getTotColor = totV(1) & "," & totV(2) & "," & totV(3) & ";" & totO(1) & "," & totO(2) & "," & totO(3) & ";" & totR(1) & "," & totR(2) & "," & totR(3)
    End If
    
End Function


Function totalColor(reference As String, ajout As String) As String
    Dim tabReference() As String
    Dim tabAjout() As String
    Dim totC(3) As Integer
    Dim i As Integer
    tabReference = Split(reference, ";")
    tabAjout = Split(ajout, ";")
    
    totalColor = ""
    For i = 0 To UBound(tabReference)
        totC(1) = val(Split(tabReference(i), ",")(0)) + val(Split(tabAjout(i), ",")(0))
        totC(2) = val(Split(tabReference(i), ",")(1)) + val(Split(tabAjout(i), ",")(1))
        totC(3) = val(Split(tabReference(i), ",")(2)) + val(Split(tabAjout(i), ",")(2))
        If totalColor = "" Then
            totalColor = totC(1) & "," & totC(2) & "," & totC(3)
        Else
            totalColor = totalColor & ";" & totC(1) & "," & totC(2) & "," & totC(3)
        End If
    Next i
End Function

Function resultatCalcul(tabReference As String)
    With ThisWorkbook.Worksheets("totalPoint")
            .Range("E2") = Split(replace(tabReference, ";", ","), ",")(0)
            .Range("H2") = Split(replace(tabReference, ";", ","), ",")(1)
            .Range("K2") = Split(replace(tabReference, ";", ","), ",")(2)
            .Range("D2") = Split(replace(tabReference, ";", ","), ",")(3)
            .Range("G2") = Split(replace(tabReference, ";", ","), ",")(4)
            .Range("J2") = Split(replace(tabReference, ";", ","), ",")(5)
            .Range("C2") = Split(replace(tabReference, ";", ","), ",")(6)
            .Range("F2") = Split(replace(tabReference, ";", ","), ",")(7)
            .Range("I2") = Split(replace(tabReference, ";", ","), ",")(8)
   End With
End Function


Function renit()
    With ThisWorkbook.Worksheets("totalPoint")
            .Range("E2") = 0
            .Range("H2") = 0
            .Range("K2") = 0
            .Range("D2") = 0
            .Range("G2") = 0
            .Range("J2") = 0
            .Range("C2") = 0
            .Range("F2") = 0
            .Range("I2") = 0
   End With
End Function






