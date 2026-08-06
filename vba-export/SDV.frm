Attribute VB_Name = "SDV"
Attribute VB_Base = "0{F9C480FC-0C4D-4958-8A11-A42B27F95027}{F74E583D-936C-426A-B383-0F281B39CB8C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub CheckBox1_Click()

Dim i As Long
    For i = 0 To Me.ListeValeur.ListCount - 1
        Me.ListeValeur.Selected(i) = CheckBox1.Value ' Select each item
    Next i

End Sub

Private Sub CommandButton1_Click()
Dim req As Object
Dim getDId As String
Dim RqOdb As Object
getDId = getDbId(Me.PL.Caption)

On Error GoTo geserreur
EventAndScreen (False)

If SelA = False Then
    MsgBox "SELECTION VIDE", vbCritical, "ODRIV"
    Exit Sub
End If


 Call Erase_All2
 Set RqOdb = db.GetOdb(val(getDId))
 Set req = db.Request("Select * from projet where ID =" & Me.PL.Caption, RqOdb)
              With ThisWorkbook.sheets("HOME")
                .Range("Project") = req.Fields(0).Value
                .Range("Fuel") = req.Fields(3).Value
                .Range("Mode") = req.Fields(11).Value
                .Range("DriveVersion") = req.Fields(10).Value
                .Range("Gears") = req.Fields(2).Value
                .Range("Area") = req.Fields(6).Value
                .Range("Prestation") = req.Fields(4).Value
                .Range("idProjects") = req.Fields(13).Value
                .Range("Targets") = "PREMIUM"
                .Range("Milestone") = req.Fields(5).Value
                .Range("Software") = req.Fields(8).Value
                .Range("C23") = req.Fields(9).Value
                .Range("H23") = req.Fields(14).Value
                ThisWorkbook.sheets("HOME").Range("AT32").Value = req.Fields(0).Value
                ThisWorkbook.sheets("HOME").Range("UNIQUEP").Value = req.Fields(13).Value
                ThisWorkbook.sheets("HOME").Range("Project").Value = req.Fields(0).Value
                If sheetExists("DATA") Then
                        Application.DisplayAlerts = False
                        ThisWorkbook.Worksheets("DATA").Delete
                        If sheetExists("GRILLE") Then ThisWorkbook.Worksheets("GRILLE").Delete
                        Application.DisplayAlerts = True
                        sheets.Add.Name = "DATA"
                 Else
                        sheets.Add.Name = "DATA"
                 End If
                 
                 Call InsertDB.chargeVal(getSDV)
                 If ThisWorkbook.Worksheets("DATA").Range("A65000").End(xlUp).row > 1 Then
                     Me.hide
                     form.hide
                     ProgressLoad
                     ProgressTitle ("Chargement Données...")
                     ThisWorkbook.Worksheets("Structure").Range("N1") = (ThisWorkbook.Worksheets("DATA").Range("A65000").End(xlUp).row) + 7
                      Call InsertDB.CVAL
                  End If
                End With
'            If Not req Is Nothing Then req.Close
            Set req = Nothing
            ThisWorkbook.sheets("TARGETS").Visible = False
            ProgressTitle ("MAJ Calcul ")
            Call delProjectEmpty
            sheets("HOME").Activate
            EventAndScreen (True)
            ProgressExit
            db.CloseSudbConn
            Unload Me
            Unload form
            MsgBox "Done", vbInformation, "ODRIV"
 
geserreur:
      If ERR.Number = 3001 Or ERR.Number = 3008 Or ERR.Number = -2147467259 Then
          MsgBox "Erreur base de données ", vbCritical, "Oliv"
           Application.EnableEvents = True
           Application.ScreenUpdating = True
           Application.Calculation = xlCalculationAutomatic
          Unload Me
           Unload form
           ProgressExit
      ElseIf ERR.Number <> 0 Then
          MsgBox ERR.description, vbCritical, "Oliv"
           Application.EnableEvents = True
           Application.ScreenUpdating = True
           Application.Calculation = xlCalculationAutomatic
           ProgressExit
           Unload Me
           Unload form
      End If
End Sub

Private Sub ListeValeur_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim req As Object
    Dim i As Long
    Dim getDId As String
    Dim RqOdb As Object
    ' Existing code for initializing form captions and database retrieval
    Me.code.Caption = form.SEL.Caption
    Me.PL.Caption = form.SELECTS.Caption
    getDId = getDbId(Me.PL.Caption)
    Set RqOdb = db.GetOdb(val(getDId))
    Set req = db.Request("Select [Sous situation de vie, Sub Event Name] from dataId WHERE UNIQUENAME=" & Me.PL.Caption & " Group By [Sous situation de vie, Sub Event Name] Order By [Sous situation de vie, Sub Event Name]", RqOdb)
    ' Adding items to ListeValeur
    If Not req Is Nothing Then
        i = 0
        While Not req.EOF
            Me.ListeValeur.AddItem req.Fields(0).Value
            req.MoveNext
            i = i + 1
        Wend
    End If
    If Not req Is Nothing Then req.Close
    Set req = Nothing
    ' Reset captions after adding items
    form.SELECTS.Caption = ""
    form.SEL.Caption = ""
End Sub

Function getSDV() As String
Dim i As Long
 getSDV = ""
For i = 0 To Me.ListeValeur.ListCount - 1
        If Me.ListeValeur.Selected(i) = True Then
               If getSDV = "" Then getSDV = Chr(34) & Me.ListeValeur.list(i) & Chr(34) Else getSDV = getSDV & ", " & Chr(34) & Me.ListeValeur.list(i) & Chr(34)
        End If
Next i
End Function

Function SelA() As Boolean
Dim i As Long
 SelA = False
For i = 0 To Me.ListeValeur.ListCount - 1
        If Me.ListeValeur.Selected(i) = True Then
            SelA = True
            Exit Function
        End If
Next i
End Function


Private Sub btnSelectAll_Click()
    Dim i As Long
    For i = 0 To Me.ListeValeur.ListCount - 1
        Me.ListeValeur.Selected(i) = True ' Select each item
    Next i
End Sub


Private Sub btnUnselectAll_Click()
    Dim i As Long
    For i = 0 To Me.ListeValeur.ListCount - 1
        Me.ListeValeur.Selected(i) = False ' Unselect each item
    Next i
End Sub



