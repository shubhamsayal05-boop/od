Attribute VB_Name = "Feuil21"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



Private Sub Worksheet_Change(ByVal Target As Range)
    Dim r As Range
    Dim i As Integer
    Dim j As Integer
    On Error GoTo Ers

    If Not Application.Intersect(Target, Columns("E:E")) Is Nothing Then
           Call DBStructure.getEntete
           If Target.Cells.Count = 1 Then
                Application.ScreenUpdating = False
                If ExistCol(replace(UCase(Target.Value), ".", "")) = False Then
                    Call Affect(replace(UCase(Target.Value), ".", ""))
                End If
                Application.ScreenUpdating = True

           Else
                Set r = Target.Cells(1, 1)
                j = (Target.Columns.Count + r.Column) - 1
                i = 1
                While r.Column <> 5 And r.Column <= j
                    i = i + 1
                    Set r = r.Offset(0, 1)
                Wend
                
                For Each r In Range(Target.Cells(1, i), Target.Cells(Target.Rows.Count, i))
                        If Len(r.Value) > 0 Then
                            If ExistCol(replace(UCase(r.Value), ".", "")) = False Then
                                Call Affect(replace(UCase(r.Value), ".", ""))
                            End If
                        End If
                Next r
           End If
    End If

Ers:
    If ERR.Number <> 0 Then
             Application.ScreenUpdating = True
              Application.EnableEvents = True
              MsgBox ERR.description, vbCritical, "ODRIVE"
    End If
End Sub

Function ExistCol(cols As String) As Boolean
  If Not sheets("DBStructure").Columns(6).Find(What:=cols, LookAt:=xlWhole) Is Nothing Then ExistCol = True Else ExistCol = False
End Function

Function Affect(cols As String)
Dim lastRow As Integer
Dim col As String
Dim table  As String

With sheets("DBStructure")
    lastRow = .Cells(.Rows.Count, 4).End(xlUp).row
    col = val(replace(.Range("D" & lastRow), "col_", "")) + 1
    If val(col) = 253 Then col = 256
    table = getTableByCol(val(col))
    
    If table <> "" Then
        col = "col_" & col
        db.Execute ("insert into entete  (IdCol, TableCol, DescriptionCol) values (" _
                        & Chr(34) & col & Chr(34) & "," _
                        & Chr(34) & table & Chr(34) & "," _
                        & Chr(34) & cols & Chr(34) & ")")
        
    End If
End With


End Function






