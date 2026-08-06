Attribute VB_Name = "Filtre"
Attribute VB_Base = "0{9F2951CE-857B-4A85-B284-27F3001CD922}{C3305B31-C723-4202-A8D6-F455992A70A3}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Sub CommandButton6_Click()
    Dim i As Integer
    Dim j As Integer
    Dim p As Integer
    Dim n As Integer
    Dim SEL    As Boolean
    Dim req As String
    Dim Parts As String
    Dim li As Variant
    req = ""
    n = 1
    SEL = False
    Parts = ""
     With form.FiltreView
                For i = 1 To .ColumnHeaders.Count
                
                    If .ColumnHeaders.Item(i).text = Me.NameProject.Value Then
                        j = .ColumnHeaders.Item(i).index
                    End If
                Next i
                
                For i = 1 To .ListItems.Count
                        .ListItems(i).ListSubItems(j - 1).text = ""
                Next i
               
                For i = 0 To Me.ListeValeur.ListCount - 1
                        If Me.ListeValeur.Selected(i) = True Then
                            If .ListItems.Count < n Then
                                Set li = .ListItems.Add(, , n)
                                li.SubItems(j - 1) = Me.ListeValeur.list(i)
                                For p = 2 To .ColumnHeaders.Count
                                    If p - 1 <> j - 1 Then
                                        li.SubItems(p - 1) = ""
                                    End If
                                Next p
                            Else
                                .ListItems(n).ListSubItems(j - 1).text = Me.ListeValeur.list(i)
                            End If
                            n = n + 1
                            SEL = True
                            
                        End If
                Next i
                If SEL = False Then
                    form.ListView1.ColumnHeaders.Item(j).text = replace(form.ListView1.ColumnHeaders.Item(j).text, "|§| ", "")
                  
                Else
              
                    If InStr(1, form.ListView1.ColumnHeaders.Item(j).text, "|§| ") = 0 Then
                        form.ListView1.ColumnHeaders.Item(j).text = "|§| " & form.ListView1.ColumnHeaders.Item(j).text
                    End If
                End If
                
                For i = 2 To .ColumnHeaders.Count
                    req = ""
                    For p = 1 To .ListItems.Count
                        If .ListItems(p).ListSubItems(i - 1).text <> "" Then
                            If req = "" Then req = "projet." & .ColumnHeaders.Item(i).text & " In (" & Chr(34) & .ListItems(p).ListSubItems(i - 1).text & Chr(34) Else _
                            req = req & ", " & Chr(34) & .ListItems(p).ListSubItems(i - 1).text & Chr(34)
                        End If
                    Next p
                    If req <> "" Then
                        req = req & ")"
                        If Parts <> "" Then req = " AND " & req
                        If Parts = "" Then Parts = req Else Parts = Parts & req
                    End If
                Next i
                form.QU.Value = Parts
                Unload Me
    End With
End Sub

Private Sub UserForm_Activate()
    Dim req As Object
    Dim i As Integer
    Dim j As Integer
    Dim CVAL As String
    CVAL = ""
   
    Set req = db.Request("Select projet." & Me.NameProject.Value & " From projet INNER JOIN projet" & db.AnneeEnCours & " ON projet.id = projet" & db.AnneeEnCours & ".code" & IIf(Len(GetqURY) > 0, " Where " & GetqURY & " And projet.code <> 'INIBDELETEPI'", " where  projet.code <> 'INIBDELETEPI'") & " Group By projet." & Me.NameProject.Value)
     
      With form.FiltreView
             If .ListItems.Count > 0 Then
                For i = 1 To .ColumnHeaders.Count
                    If .ColumnHeaders.Item(i).text = Me.NameProject.Value Then
                        j = .ColumnHeaders.Item(i).index
                    End If
                Next i
                For i = 1 To .ListItems.Count
                      If CVAL = "" Then CVAL = .ListItems(i).ListSubItems(j - 1).text Else CVAL = CVAL & ";" & .ListItems(i).ListSubItems(j - 1).text
                Next i
                CVAL = ";" & CVAL & ";"
            End If
            
            If Not req Is Nothing Then
                Me.ListeValeur.Clear
                While Not req.EOF
'                    If Me.NameProject.Value <> "SOFTWARE" Then
                    If Len(req.Fields(0).Value) > 0 Or IsNull(req.Fields(0).Value) = False Then
                        Me.ListeValeur.AddItem req.Fields(0).Value
                  
'                    Else
'                        Me.ListeValeur.AddItem db.GetValue("SELECT software FROM milestone WHERE id= " & Req.Fields(0).Value & " ")
'                    End If
                        If InStr(1, CVAL, ";" & Me.ListeValeur.list(Me.ListeValeur.ListCount - 1) & ";") <> 0 Then
                            Me.ListeValeur.Selected(Me.ListeValeur.ListCount - 1) = True
                        End If
                    
                     End If
                    req.MoveNext
                Wend
            End If
             If Not req Is Nothing Then req.Close
             Set req = Nothing
        End With
End Sub



Function GetqURY()
    Dim i As Integer
    Dim p As Integer
    Dim req As String
    Dim Parts As String
    i = 2
     With form.ListView1
        While replace(.ColumnHeaders.Item(i).text, "|§| ", "") <> tRANsT(Me.NameProject.Value)
            If InStr(1, .ColumnHeaders.Item(i).text, "|§|") <> 0 Then
                 req = ""
                 For p = 1 To form.FiltreView.ListItems.Count
                    If form.FiltreView.ListItems(p).ListSubItems(i - 1).text <> "" Then
                        If req = "" Then req = form.FiltreView.ColumnHeaders.Item(i).text & " In (" & Chr(34) & form.FiltreView.ListItems(p).ListSubItems(i - 1).text & Chr(34) Else _
                        req = req & ", " & Chr(34) & form.FiltreView.ListItems(p).ListSubItems(i - 1).text & Chr(34)
                    End If
                Next p
                If req <> "" Then
                      req = req & ")"
                      If Parts <> "" Then req = " AND " & req
                      If Parts = "" Then Parts = req Else Parts = Parts & req
                End If
            End If
            i = i + 1
            
        Wend
        GetqURY = Parts
    End With
End Function

Function tRANsT(Vas As String)
    
    
    If Vas = "CODE" Then
         tRANsT = "NAME/ CODE"
    ElseIf Vas = "GEARS" Then
          tRANsT = "GEAR"
    ElseIf Vas = "MILESTONE" Then
          tRANsT = "ODRIV MILESTONE"
    ElseIf Vas = "AERA" Then
          tRANsT = "AREA"
    ElseIf Vas = "SOFTWARE" Then
          tRANsT = "SOFTWARE MILESTONE"
    ElseIf Vas = "TARGET_VEHICLE" Then
          tRANsT = "TARGET VEHICLE"
    Else
          tRANsT = Vas
    End If
   
   
End Function




