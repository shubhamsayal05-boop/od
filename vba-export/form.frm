Attribute VB_Name = "form"
Attribute VB_Base = "0{FD30A9D8-05CE-4B8C-A285-5F0DD037B952}{016720C4-5933-4164-8A41-E61198BDB9BE}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub CANC_Click()
    Me.CANC.Visible = False
    Me.ListView1.CheckBoxes = False
    Me.ListView1.MultiSelect = False
    Me.CommandButton6.Visible = True
    Me.DELP.Caption = "DELETE"
    SELECTS.Caption = ""
    SEL.Caption = ""
End Sub

Private Sub CommandButton6_Click()

    If Len(Me.SELECTS.Caption) > 0 Then
        SDV.Show
    Else
        MsgBox "Sélectionner d'abord un projet", vbCritical, "ODRIV"
    End If
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim j
    If ListView1.ListItems(Item.index).Checked = True Then
        SELECTS.Caption = IIf(SELECTS.Caption = "", ListView1.ListItems(Item.index).text, SELECTS.Caption & "," & ListView1.ListItems(Item.index).text)
    Else
      If InStr(1, "," & SELECTS.Caption & ",", "," & ListView1.ListItems(Item.index) & ",") <> 0 Then
         SELECTS.Caption = replace(SELECTS.Caption, ListView1.ListItems(Item.index), "")
          For j = 1 To Len(SELECTS.Caption)
                SELECTS.Caption = replace(SELECTS.Caption, ",,", ",")
         Next j
         If Left(SELECTS.Caption, 1) = "," Then SELECTS.Caption = Right(SELECTS.Caption, Len(SELECTS.Caption) - 1)
         If Right(SELECTS.Caption, 1) = "," Then SELECTS.Caption = Left(SELECTS.Caption, Len(SELECTS.Caption) - 1)
      End If
    End If
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
    Dim index As Long
  
    If Me.CANC.Visible = True Then Exit Sub
    For index = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(index).Selected = True Then
                SEL.Caption = ListView1.ListItems(index).ListSubItems(2).text & "_" & ListView1.ListItems(index).ListSubItems(11).text
                SELECTS.Caption = ListView1.ListItems(index).text
                Exit Sub
        End If
   Next index
End Sub

Private Sub QU_Change()
        InitialiseList
End Sub

Private Sub quitter_Click()
    Unload Me
End Sub

Private Sub DELP_Click()
   Dim index
   Dim getAll As String
   Dim stReq As Object
   Dim talbleId
   Dim getsubid As String, cGetArray() As String, cGetArray2() As String
   Dim SelectDBiD(4) As String
   Dim j, n
   Dim p As Integer
   Dim conn As Object, dbS As Object
   Dim idc As String
   Dim RqOdb As Object
   
    SelectDBiD(1) = ""
    SelectDBiD(2) = ""
    SelectDBiD(3) = ""
    SelectDBiD(4) = ""
      
        If Me.CANC.Visible = False Then
            Me.ListView1.CheckBoxes = True
            Me.ListView1.MultiSelect = True
            Me.CommandButton6.Visible = False
            Me.DELP.Caption = "OK"
            Me.CANC.Visible = True
            SELECTS.Caption = ""
            SEL.Caption = ""
            For index = 1 To ListView1.ListItems.Count
                ListView1.ListItems(index).Checked = True
                ListView1.ListItems(index).Checked = False
            Next index
            Exit Sub
        End If
        
        getsubid = ""
        If Len(SELECTS.Caption) > 0 Then
            If MsgBox("Are you Sure ?", vbCritical + vbYesNo) = vbYes Then
                 If InputBox("Tapez OUI En Majuscule Pour Confirmer", "Confirmation") = "OUI" Then
                    'INIB SUPPRESSION
                    If InStr(1, SELECTS.Caption, ",") <> 0 Then
                        cGetArray = Split(SELECTS.Caption, ",")
                    Else
                       ReDim cGetArray(0)
                       cGetArray(0) = SELECTS.Caption
                    End If
                    
                    idc = getDbId(cGetArray(n))
                    
                    If idc = "" Then
                        MsgBox "Id non disponible dans la base", vbCritical, "ODRIV"
                        Exit Sub
                    End If
                    
                    For n = 0 To UBound(cGetArray)
                        idc = getDbId(cGetArray(n))
                       'For p = 1 To 4
                               p = 1
'                               If InStr(1, idc, p) <> 0 Then
                                    Set RqOdb = db.GetOdb(CInt(idc))
'                                    Set stReq = db.Request("SELECT n° FROM dataid where uNIQUEnAME in (" & cGetArray(n) & ") order by N°", RqOdb)
                                    Set stReq = db.Request("SELECT n° FROM dataid where UniqueName in (" & cGetArray(n) & ") order by N°", RqOdb)
                                    
                                    If Not stReq Is Nothing Then
'                                             stReq.movefirst
'                                             Debug.Print (stReq(0))
                                             talbleId = stReq.getrows

                                             If SelectDBiD(p) = "" Then SelectDBiD(p) = CStr(talbleId(0, 0)) & ":" & CStr(talbleId(0, UBound(talbleId, 2))) _
                                             Else: SelectDBiD(p) = SelectDBiD(p) & "#" & CStr(talbleId(0, 0)) & ":" & CStr(talbleId(0, UBound(talbleId, 2)))

                                            ' SelectDBiD(p) = getsubid
                                    Else
                                        If SelectDBiD(p) = "" Then SelectDBiD(p) = ":" Else: SelectDBiD(p) = SelectDBiD(p) & "#" & ":"
                                    End If
'                                End If
                      'Next p
                    Next n
                    
                    If SelectDBiD(1) = "" And SelectDBiD(2) = "" And SelectDBiD(3) = "" And SelectDBiD(4) = "" Then
                                MsgBox "Id has no SDV", vbInformation, "ODRIV"
                                For n = 0 To UBound(cGetArray)
                                    idc = getDbId(CInt(cGetArray(n)))
                                    Set conn = CreateObject("DAO.DBEngine.120")
                                    Set dbS = conn.OpenDatabase(GetYearDbPath("_OdrivDB_" & CInt(idc)))
                                    Call ExecuteAccessQuery("Delete From dataId where UniqueName in (" & CInt(cGetArray(n)) & ")", dbS)
                                    Call ExecuteAccessQuery("DELETE * FROM projet WHERE ID in (" & CInt(cGetArray(n)) & ")", dbS)
                                    dbS.Close
                                    Set dbS = Nothing
                                Next
                                db.Execute ("delete FROM PROJET WHERE ID in (" & SELECTS.Caption & ")")
                                db.Execute ("delete FROM PROJET" & db.AnneeEnCours & " WHERE code in (" & SELECTS.Caption & ")")
                                Call InitialiseList
                                
                                SELECTS.Caption = ""
                                SEL.Caption = ""
                                
                                Me.CANC.Visible = False
                                Me.ListView1.CheckBoxes = False
                                Me.ListView1.MultiSelect = False
                                Me.CommandButton6.Visible = True
                                Me.DELP.Caption = "DELETE"
                                MsgBox "Delete Successful", vbInformation, "ODRIV"
                                Exit Sub
                    End If

                    If Not stReq Is Nothing Then stReq.Close
                     Set stReq = Nothing
                     
'                     For p = 1 To 4
                       p = 1
                       If Len(SelectDBiD(p)) > 0 Then
                                Set conn = CreateObject("DAO.DBEngine.120")
'                                Set dbS = conn.OpenDatabase(ThisWorkbook.Worksheets("cfg").Range("B1").Value & "\_OdrivDB_" & p & ".accdb")
                                If InStr(1, SelectDBiD(p), "#") <> 0 Then
                                    cGetArray2 = Split(SelectDBiD(p), "#")
                                Else
                                   ReDim cGetArray2(0)
                                   cGetArray2(0) = SelectDBiD(p)
                                End If
                                
                                For n = 0 To UBound(cGetArray)
                                    If Split(cGetArray2(n), ":")(0) <> vbNullString And Split(cGetArray2(n), ":")(1) <> vbNullString Then
                                        idc = getDbId(CInt(cGetArray(n)))
                                        Set dbS = conn.OpenDatabase(GetYearDbPath("_OdrivDB_" & CInt(idc)))
                                        Call ExecuteAccessQuery("DELETE idData FROM dataSub1 WHERE idData>=" & Split(cGetArray2(n), ":")(0) & " And idData<=" & Split(cGetArray2(n), ":")(1), dbS)
                                        Call ExecuteAccessQuery("DELETE idData FROM dataSub2 WHERE idData>=" & Split(cGetArray2(n), ":")(0) & " And idData<=" & Split(cGetArray2(n), ":")(1), dbS)
                                        Call ExecuteAccessQuery("DELETE idData FROM dataSub3 WHERE idData>=" & Split(cGetArray2(n), ":")(0) & " And idData<=" & Split(cGetArray2(n), ":")(1), dbS)
                                        Call ExecuteAccessQuery("Delete From dataId where UniqueName in (" & CInt(cGetArray(n)) & ")", dbS)
                                        Call ExecuteAccessQuery("DELETE * FROM projet WHERE ID in (" & CInt(cGetArray(n)) & ")", dbS)
    '                                If UBound(cGetArray) = 0 Then
    '                                    db.Execute ("DELETE idData FROM dataSub1 WHERE idData>=" & Split(cGetArray(n), ":")(0) & " And idData<=" & Split(cGetArray(n), ":")(0))
    '                                    db.Execute ("DELETE idData FROM dataSub2 WHERE idData>=" & Split(cGetArray(n), ":")(0) & " And idData<=" & Split(cGetArray(n), ":")(0))
    '                                    db.Execute ("DELETE idData FROM dataSub3 WHERE idData>=" & Split(cGetArray(n), ":")(0) & " And idData<=" & Split(cGetArray(n), ":")(0))
    '                                Else
    '                                    db.Execute ("DELETE idData FROM dataSub1 WHERE idData>=" & Split(cGetArray(n), ":")(0) & " And idData<=" & Split(cGetArray(n), ":")(1))
    '                                    db.Execute ("DELETE idData FROM dataSub2 WHERE idData>=" & Split(cGetArray(n), ":")(0) & " And idData<=" & Split(cGetArray(n), ":")(1))
    '                                    db.Execute ("DELETE idData FROM dataSub3 WHERE idData>=" & Split(cGetArray(n), ":")(0) & " And idData<=" & Split(cGetArray(n), ":")(1))
    '                                End If
                                        dbS.Close
                                        Set dbS = Nothing
                                    End If
                                Next n
                                For n = 0 To UBound(cGetArray)
                                    idc = getDbId(CInt(cGetArray(n)))
                                    Set conn = CreateObject("DAO.DBEngine.120")
                                    Set dbS = conn.OpenDatabase(GetYearDbPath("_OdrivDB_" & CInt(idc)))
                                    Call ExecuteAccessQuery("Delete From dataId where UniqueName in (" & CInt(cGetArray(n)) & ")", dbS)
                                    Call ExecuteAccessQuery("DELETE * FROM projet WHERE ID in (" & CInt(cGetArray(n)) & ")", dbS)
                                    dbS.Close
                                    Set dbS = Nothing
                                Next
'                                db.Execute ("Delete From DATAID where uNIQUEnAME in (" & SELECTS.Caption & ")")
'                                db.Execute ("DELETE * FROM PROJET WHERE ID in (" & SELECTS.Caption & ")")
                                
                         End If
'                    Next p
                    
'                    db.Execute ("Delete From DATAID where N° in (" & SELECTS.Caption & ")")
'                    db.Execute ("Delete From DATAID where N° in (" & SELECTS.Caption & ")")

                    db.Execute ("delete FROM PROJET WHERE ID in (" & SELECTS.Caption & ")")
                    db.Execute ("delete FROM PROJET" & db.AnneeEnCours & " WHERE code in (" & SELECTS.Caption & ")")
'''''''''''''''                     db.Execute ("UPDATE PROJET SET CODE='INIBDELETEPI' WHERE ID in (" & SELECTS.Caption & ")")
                   
'                     dbS.Close
'                     Set dbS = Nothing
                     
                    Call InitialiseList
                    
                    MsgBox "Delete Successful", vbInformation, "ODRIV"
                    SELECTS.Caption = ""
                    SEL.Caption = ""
                    
                    Me.CANC.Visible = False
                    Me.ListView1.CheckBoxes = False
                    Me.ListView1.MultiSelect = False
                    Me.CommandButton6.Visible = True
                    Me.DELP.Caption = "DELETE"
                Else
                     Me.CANC.Visible = False
                     Me.ListView1.CheckBoxes = False
                     Me.ListView1.MultiSelect = False
                     Me.CommandButton6.Visible = True
                     Me.DELP.Caption = "DELETE"
                     SELECTS.Caption = ""
                     SEL.Caption = ""
                     MsgBox "SUPPRESSION ANNULEE", vbCritical, "ODRIV"
                     
                     
                 End If
            Else
'                SELECTS.Caption = ""
'                SEL.Caption = ""
            End If
        End If
End Sub

Private Sub ListView1_Click()
    
    Dim index As Long
  
    If Me.CANC.Visible = True Then Exit Sub
    For index = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(index).Selected = True Then
                SEL.Caption = ListView1.ListItems(index).ListSubItems(2).text & "_" & ListView1.ListItems(index).ListSubItems(11).text
                SELECTS.Caption = ListView1.ListItems(index).text
                Exit Sub
        End If
   Next index
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim colonneNames As String
    Dim CHeader As String
    
    CHeader = replace(ColumnHeader.text, "|§| ", "")
    If CHeader = "ID" Then Exit Sub
    If CHeader = "NAME/ CODE" Then
         colonneNames = "CODE"
    ElseIf CHeader = "GEAR" Then
          colonneNames = "GEARS"
    ElseIf CHeader = "ODRIV MILESTONE" Then
          colonneNames = "MILESTONE"
    ElseIf CHeader = "AREA" Then
          colonneNames = "AERA"
    ElseIf CHeader = "SOFTWARE MILESTONE" Then
          colonneNames = "SOFTWARE"
    ElseIf CHeader = "TARGET VEHICLE" Then
          colonneNames = "TARGET_VEHICLE"
    Else
          colonneNames = CHeader
    End If
   
   Filtre.NameProject.Value = colonneNames
   Filtre.Show
End Sub

Private Sub InitialiseList(Optional trou As Boolean)
    Dim k As Long
    Dim dat_proj As Variant
    Dim i As Long

         dat_proj = db.GetListValue("SELECT  projet.code, projet.droopy, projet.gears, projet.energy, projet.MODE, projet.milestone, projet.aera, projet.target, projet.software, projet.target_vehicle, projet.Version, projet.Mode, projet.Uniquename, projet.ID, projet.NbGear FROM projet INNER JOIN projet" & db.AnneeEnCours & " ON projet.id = projet" & db.AnneeEnCours & ".code" & IIf(Len(Me.QU.Value) > 0, " WHERE " & Me.QU.Value & " And projet.code <> 'INIBDELETEPI' ORDER BY projet.id", " where projet.code <> 'INIBDELETEPI' ORDER BY projet.id"))
        If Not isEmpty(dat_proj) Then
        
            Dim li As Object
            Dim NB As Integer
            Dim d_id As String, d_version As String, d_droop As String, d_nam As String, d_gea As String, d_ener As String, d_prio As String, d_odri As String, d_are As String, d_tar As String, d_soft As String, d_typ As String, d_vehi As String
             
'''            If trou <> True Then
'''                ComboProject.AddItem ""
'''            End If
            


           
              ListView1.ListItems.Clear
               
               With ListView1
                   
                    'Définit le nombre de colonnes et Entêtes
                  If Len(Me.QU.Value) = 0 Then
                   With .ColumnHeaders
                       'Supprime les anciens entêtes
                       .Clear
                       'Ajoute 3 colonnes en spécifiant le nom de l'entête
                       'et la largeur des colonnes
                       .Add , , "ID"
                       .Add , , "DROOPY"
                       .Add , , "NAME/ CODE", 100
                       .Add , , "GEAR"
                       .Add , , "ENERGY"
                       .Add , , "MODE", 100
                       .Add , , "ODRIV MILESTONE", 100
                       .Add , , "AREA"
                       .Add , , "TARGET"
                       .Add , , "SOFTWARE MILESTONE", 100
                       .Add , , "TARGET VEHICLE"
                       .Add , , "VERSION"
                   End With
                End If
                   .View = lvwReport   'affichage en mode Rapport
                 .Gridlines = True   'affichage d'un quadrillage
                 .FullRowSelect = True   'Sélection des lignes comlètes
                 .LabelEdit = lvwManual  'desactive edition du listview
                  
               End With
            
        
        
    
              For i = 0 To UBound(dat_proj, 2)
                 
                 NB = NB + 1
                 d_id = CStr(NB)
                 

                    

                            d_droop = dat_proj(1, i)
                            d_nam = dat_proj(0, i)
                            d_gea = dat_proj(2, i)
                            d_ener = dat_proj(3, i)
                            d_prio = IIf(IsNull(dat_proj(4, i)) Or Len(dat_proj(4, i)) = 0, "", dat_proj(4, i))
                            d_odri = dat_proj(5, i)
                            d_are = dat_proj(6, i)
                            d_tar = dat_proj(7, i)
                           d_vehi = dat_proj(9, i)
                           d_version = dat_proj(10, i)
                           d_id = dat_proj(13, i)
                           d_soft = dat_proj(8, i)

'                     ElseIf K = 8 Then
'
'                        sof = db.GetValue("SELECT software FROM milestone WHERE id= " & dat_proj(K, I) & " ")
'                         d_soft = sof
'                     End If
                     
 
                 
                 Set li = ListView1.ListItems.Add(, , d_id)
                 li.SubItems(1) = d_droop
                 li.SubItems(2) = d_nam
                 li.SubItems(3) = d_gea
                 li.SubItems(4) = d_ener
                 li.SubItems(5) = d_prio
                 li.SubItems(6) = d_odri
                 li.SubItems(7) = d_are
                 li.SubItems(8) = d_tar
                 li.SubItems(9) = d_soft
                 li.SubItems(10) = d_vehi
                 li.SubItems(11) = d_version
                
'                If StrComp(d_typ, "Draft", vbTextCompare) = 0 Then
'                    'li.Bold = True
'                    li.ForeColor = vbRed
'
'                    ListView1.ListItems(I + 1).ListSubItems(1).ForeColor = vbRed
'                    ListView1.ListItems(I + 1).ListSubItems(2).ForeColor = vbRed
'                    ListView1.ListItems(I + 1).ListSubItems(3).ForeColor = vbRed
'                    ListView1.ListItems(I + 1).ListSubItems(4).ForeColor = vbRed
'                    ListView1.ListItems(I + 1).ListSubItems(5).ForeColor = vbRed
'                    ListView1.ListItems(I + 1).ListSubItems(6).ForeColor = vbRed
'                    ListView1.ListItems(I + 1).ListSubItems(7).ForeColor = vbRed
'                    ListView1.ListItems(I + 1).ListSubItems(8).ForeColor = vbRed
'                    ListView1.ListItems(I + 1).ListSubItems(9).ForeColor = vbRed
'                    ListView1.ListItems(I + 1).ListSubItems(10).ForeColor = vbRed
'                    ListView1.ListItems(I + 1).ListSubItems(11).ForeColor = vbRed
'                 Else
'                    li.Bold = True
'                    li.ForeColor = vbBlue
'
                    ListView1.ListItems(i + 1).ListSubItems(1).Bold = True
                    ListView1.ListItems(i + 1).ListSubItems(1).ForeColor = vbBlue
                    ListView1.ListItems(i + 1).ListSubItems(2).Bold = True
                    ListView1.ListItems(i + 1).ListSubItems(2).ForeColor = vbBlue
                    ListView1.ListItems(i + 1).ListSubItems(3).Bold = True
                    ListView1.ListItems(i + 1).ListSubItems(3).ForeColor = vbBlue
                    ListView1.ListItems(i + 1).ListSubItems(4).Bold = True
                    ListView1.ListItems(i + 1).ListSubItems(4).ForeColor = vbBlue
                    ListView1.ListItems(i + 1).ListSubItems(5).Bold = True
                    ListView1.ListItems(i + 1).ListSubItems(5).ForeColor = vbBlue
                    ListView1.ListItems(i + 1).ListSubItems(6).Bold = True
                    ListView1.ListItems(i + 1).ListSubItems(6).ForeColor = vbBlue
                    ListView1.ListItems(i + 1).ListSubItems(7).Bold = True
                    ListView1.ListItems(i + 1).ListSubItems(7).ForeColor = vbBlue
                    ListView1.ListItems(i + 1).ListSubItems(8).Bold = True
                    ListView1.ListItems(i + 1).ListSubItems(8).ForeColor = vbBlue
                    ListView1.ListItems(i + 1).ListSubItems(9).Bold = True
                    ListView1.ListItems(i + 1).ListSubItems(9).ForeColor = vbBlue
                    ListView1.ListItems(i + 1).ListSubItems(10).Bold = True
                    ListView1.ListItems(i + 1).ListSubItems(10).ForeColor = vbBlue
                    ListView1.ListItems(i + 1).ListSubItems(11).Bold = True
                    ListView1.ListItems(i + 1).ListSubItems(11).ForeColor = vbBlue
'                 End If
'
             Next i
             
'             NameProject.Value = ""
        Else
            ListView1.ListItems.Clear
        End If
        
  
End Sub

Private Sub UserForm_Initialize()
Dim i

Call InitialiseList
Call CFiltre
Me.CANC.Visible = False
End Sub

Function CFiltre()
       
            
            With FiltreView
                 
                  'Définit le nombre de colonnes et Entêtes
                 With .ColumnHeaders
                     'Supprime les anciens entêtes
                     .Clear
                     'Ajoute 3 colonnes en spécifiant le nom de l'entête
                     'et la largeur des colonnes
                     .Add , , "ID"
                     .Add , , "DROOPY"
                     .Add , , "CODE", 100
                     .Add , , "GEARS"
                     .Add , , "ENERGY"
                     .Add , , "MODE", 100
                     .Add , , "MILESTONE", 100
                     .Add , , "AERA"
                     .Add , , "TARGET"
                     .Add , , "SOFTWARE", 100
                     .Add , , "TARGET_VEHICLE"
                     .Add , , "VERSION"
                     
                 End With
             
                 .View = lvwReport   'affichage en mode Rapport
               .Gridlines = True   'affichage d'un quadrillage
               .FullRowSelect = True   'Sélection des lignes comlètes
               .LabelEdit = lvwManual  'desactive edition du listview
                
               
                
             End With
              
End Function




Private Sub txtSearchID_Change()

    If Len(Me.txtSearchID.Value) > 0 Then
        Me.QU.Value = "projet.ID LIKE '%" & Me.txtSearchID.Value & "%'"
    Else
        Me.QU.Value = ""
    End If

    InitialiseList True

End Sub










