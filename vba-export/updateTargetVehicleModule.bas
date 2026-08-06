Attribute VB_Name = "updateTargetVehicleModule"
Option Explicit
Private dict As Object

Function updateTargetVehicle()
    Dim l As Long, j As Long
    Dim indexesArray As String
    getTargetFromVHDict
     l = getLastRowRating
    With ThisWorkbook.Worksheets("RATING")
        For j = 23 To l
            If .Rows(j).Hidden = False And .Cells(j, 4) <> "" Then
               indexesArray = .Range("D" & j) & "-" & _
                                     .Range("M" & j) & "-" & _
                                     .Cells(j, .Rows("21:22").Find(What:="Responsiveness Index", LookAt:=xlWhole).Column).Value
               Call checkAndUpdateTarget(indexesArray)
            End If
        Next j
         indexesArray = "Rate of low points" & "-" & _
                                     .Range("AM12") & "-" & _
                                    .Range("AM18")
         Call checkAndUpdateTarget(indexesArray)
    End With
End Function

Function getTargetFromVHDict()
    Dim lrow As Long
    Dim key As String
    Dim ws As Worksheet
    Dim i As Integer
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.sheets("TARGET VEHICLE")
    lrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    For i = 2 To lrow
        key = ws.Cells(i, 1).Value & "-" & ws.Cells(i, 2).Value & "-" & ws.Cells(i, 3).Value & "-" & ws.Cells(i, 4).Value
        If Not dict.Exists(key) Then
            dict.Add key, i
        End If
    Next i
End Function
Function checkAndUpdateTarget(valToUpdate As String)
    Dim keyToFind As String
    Dim homeParam(3) As String
    Dim rowIndex As Long
    
    With ThisWorkbook.Worksheets("HOME")
       homeParam(1) = .Range("DriveVersion")
       homeParam(2) = .Range("C23")
       homeParam(3) = .Range("Mode")
    End With
    
    keyToFind = Split(valToUpdate, "-")(0) & "-" & homeParam(1) & "-" & homeParam(2) & "-" & homeParam(3)
    If dict.Exists(keyToFind) Then
        rowIndex = dict(keyToFind)
        ThisWorkbook.sheets("TARGET VEHICLE").Cells(rowIndex, 5) = Split(valToUpdate, "-")(1)
        ThisWorkbook.sheets("TARGET VEHICLE").Cells(rowIndex, 6) = Split(valToUpdate, "-")(2)
    Else
        Call insertRow(valToUpdate)
        Call getTargetFromVHDict
    End If
End Function

Sub insertRow(valToInsert As String)
    Dim l As Long
    '
    With ThisWorkbook.sheets("TARGET VEHICLE")
         l = .Cells(Rows.Count, 1).End(xlUp).row + 1
         .Range("A" & l - 1 & ":F" & l - 1).Copy Destination:=.Range("A" & l)
         .Range("A" & l & ":F" & l).ClearContents
          
          .Range("A" & l) = Split(valToInsert, "-")(0)
          .Range("B" & l) = ThisWorkbook.sheets("HOME").Range("DriveVersion")
          .Range("C" & l) = ThisWorkbook.sheets("HOME").Range("C23")
          .Range("D" & l) = ThisWorkbook.sheets("HOME").Range("Mode")
          .Range("E" & l) = Split(valToInsert, "-")(1)
          .Range("F" & l) = Split(valToInsert, "-")(2)
    End With
    
End Sub






