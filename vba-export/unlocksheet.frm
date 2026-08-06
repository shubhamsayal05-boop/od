Attribute VB_Name = "unlocksheet"
Attribute VB_Base = "0{95E82670-7A08-45B3-8763-02C1C75EC74D}{08E34DE8-96E8-4306-A88F-F92676DDDF72}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub CommandButton1_Click()
    
        
            If UCase(Me.TextBox2.Value) = "UNLOCK" Then
                If SelectionActuel(-1) = False Then
                    MsgBox "COCHER POUR SELECTIONNER", vbCritical, "ODRIV"
                Else
                    Unload Me
                End If
            ElseIf UCase(Me.TextBox2.Value) = "LOCK" Then
                If SelectionActuel(2) = False Then
                    MsgBox "COCHER POUR SELECTIONNER", vbCritical, "ODRIV"
                Else
                    Unload Me
                End If
            Else
                MsgBox "MOT DE PASSE INCORRECT", vbCritical, "ODRIV"
            End If
      
End Sub



Private Sub ListBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
       
        
End Sub



Private Sub ListeValeurS_ItemClick(ByVal Item As MSComctlLib.ListItem)
        Me.ListeValeurS.ListItems(Item.index).Selected = False
End Sub

Private Sub UserForm_Initialize()
        With Me.ListeValeurS
               .View = lvwReport   'affichage en mode Rapport
               .Gridlines = True   'affichage d'un quadrillage
               .FullRowSelect = False   'Sélection des lignes comlètes
               .LabelEdit = lvwManual  'desactive edition du listview
               .MultiSelect = True
               .HideSelection = True
               .HoverSelection = False
               .CheckBoxes = True
        End With
       
        Call loadME
       
End Sub

Function loadME()
    Dim tables(13) As String
    Dim i As Integer
    
    tables(1) = "CONFIGURATIONS SEETINGS"
    tables(2) = "SETTINGS"
    tables(3) = "structure"
    tables(4) = "POWERTRAIN"
    tables(5) = "CONFIGURATIONS"
'    TabLes(6) = "COVERAGE Rate"
    tables(6) = "Target vehicle"
    tables(7) = "TARGETS"
    tables(8) = "Calculs"
    tables(9) = "Graph_status"
    tables(10) = "DEFINITION SDV"
    tables(11) = "PARAMETRES GRAPH"
    tables(12) = "ENTETE_COLONNE"
    tables(13) = "SDV MANAGER"
    
    For i = 1 To UBound(tables)
             If sheets(tables(i)).Visible = 2 Or sheets(tables(i)).Visible = -1 Then ' or sheets(tables(i)).Visible = -1
                   Me.ListeValeurS.ListItems.Add , , UCase(tables(i))
             End If
    Next i

   
End Function
Function SelectionActuel(state) As Boolean
Dim i As Long
SelectionActuel = False
 
For i = 1 To Me.ListeValeurS.ListItems.Count
        If Me.ListeValeurS.ListItems(i).Checked = True Then
             sheets(Me.ListeValeurS.ListItems(i).text).Visible = state
             SelectionActuel = True
        End If
Next i
End Function


