Attribute VB_Name = "OrdreS"
Attribute VB_Base = "0{B742349F-DE96-44E0-9F39-6F71546A6102}{A08AC884-94E9-49E4-882B-0DF8AA170300}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Option Explicit
Private uncheck As String
Private checkList As String
Private cOrder As Object
Private chargeD As String
Private saveSelNd As String

Private Sub AjAnnuler_Click()
   Me.EtatAjout = ""
   chargeD = "x"
   UserForm_Activate
End Sub

Private Sub addItems_Click()
        addItemMod (True)
End Sub
Function addItemMod(IsAdd As Boolean)
        Dim notIsAdd As Boolean
        
        If IsAdd = True Then notIsAdd = False Else notIsAdd = True
        If IsAdd = True Then Me.EtatAjout.Caption = "ADD"
        Me.OrdreView.CheckBoxes = notIsAdd
        Me.Label12.Visible = notIsAdd
        Me.Label15.Visible = notIsAdd
        Me.CommandButton1.Visible = notIsAdd
        Me.btn_ok.Visible = notIsAdd
        Me.Annuler.Visible = False
        Me.chapitre.Visible = notIsAdd
        Me.fonction.Visible = notIsAdd
        Me.COPIE.Visible = notIsAdd
        Me.INFO.Visible = notIsAdd
        Me.addItems.Visible = notIsAdd
        Me.TitreType.Visible = IsAdd
        Me.TitreNom.Visible = IsAdd
        Me.TitrePosition.Visible = IsAdd
        Me.AjChapitre.Visible = IsAdd
        'Me.AjFonction.Visible = IsAdd
        Me.AjNom.Visible = IsAdd
        Me.AjPosition.Visible = IsAdd
        Me.AjOK.Visible = IsAdd
        Me.AjAnnuler.Visible = IsAdd
                
        If IsAdd = True Then
        
                Me.Label13.Top = 438
                Me.OrdreView.Top = 126
                Me.TitreType.Top = 30
                Me.TitreNom.Top = 78
                Me.TitrePosition.Top = 78
                Me.AjChapitre.Top = 54
                Me.AjFonction.Top = 54
                Me.AjNom.Top = 99
                Me.AjPosition.Top = 99
                Me.AjOK.Top = 42
                Me.AjAnnuler.Top = 42
                Me.AjChapitre.Value = True
                Me.AjFonction.Value = False
                Me.AjPosition.Enabled = False
                Me.AjPosition.Locked = True
                Me.AjNom.Locked = False
                Me.AjPosition.Value = "Selectionner Ci dessous..."
                Me.TitrePosition.Caption = "Position "
                
                
       Else
                Me.Label13.Top = 414
                Me.OrdreView.Top = 90
                Me.EtatAjout.Caption = ""
                Me.AjNom.Value = ""
                Me.AjPosition.Value = "Selectionner Ci dessous..."
                Me.AjChapitre.Value = False
                Me.AjFonction.Value = False
       End If
       
       If Me.OrdreView.nodes.Count > 0 Then Me.OrdreView.nodes(1).Selected = False
'
End Function

Private Sub AjOK_Click()
        Dim Schap As String
        If Me.AjChapitre.Value = False And Me.AjFonction.Value = False Then
            MsgBox "Choisir Group", vbCritical, "Bilique"
        ElseIf Len(Me.AjNom.Value) = 0 Then
            MsgBox "Nom Vide", vbCritical, "Bilique"
        ElseIf Me.AjPosition.Value = "Selectionner Ci dessous..." Then
            MsgBox "Selectionner Position", vbCritical, "Bilique"
        ElseIf nodeExist(Me.AjNom.Value) = True Then
             MsgBox "Existe Deja", vbCritical, "Bilique"
        Else
            
            If Me.AjFonction.Value = True Then
                 moveCheckedValue ("AddFonction")
            Else
                 moveCheckedValue ("AddChapitre")
            End If
            Me.addItems.Visible = False
           If Me.OrdreView.nodes.Count > 0 Then Me.OrdreView.nodes(1).Selected = False
        End If
End Sub



Private Sub Annuler_Click()
    Me.COPIE.Caption = "Clique droit Pour Couper"
    Me.OrdreView.CheckBoxes = True
    defautColor
    Me.chapitre.Enabled = True
    Me.addItems.Visible = True
    Me.fonction.Enabled = True
    checkList = ""
    Me.Annuler.Visible = False
    If Me.OrdreView.nodes(Me.OrdreView.nodes.Count).text = "Coller Ici" Then
         Me.OrdreView.nodes.Remove Me.OrdreView.nodes.Count
    End If
    
End Sub

Private Sub btn_ok_Click()
    On Error GoTo Ers
    
    Application.ScreenUpdating = False
    Call validateAll
    Application.ScreenUpdating = True
    
    Unload Me
    
Ers:
    If ERR.Number <> 0 Then
        MsgBox ERR.description, vbCritical, "BILIQUE"
        Application.ScreenUpdating = True
        Application.EnableEvents = True
    End If
End Sub

Private Sub chapitre_Change()
        unckeckALL
End Sub
Private Sub AjChapitre_Change()
       If Me.AjChapitre.Value = True Then
            chargeChapFonct ("c")
        ElseIf Me.AjFonction.Value = True Then
            chargeChapFonct ("f")
        End If
End Sub
Private Sub AjFonction_Change()

        
        If Me.AjFonction.Value = True Then
            chargeChapFonct ("f")
        ElseIf Me.AjChapitre.Value = True Then
            chargeChapFonct ("c")
        End If
End Sub
Private Sub Fonction_Change()
        unckeckALL
End Sub
Function MaskT(v As Integer)
        If v = 0 Then
            Me.OrdreView.Visible = False
        Else
            Me.OrdreView.Visible = True
            makeBold
        End If
End Function
Function chargeChapFonct(c As String, Optional maske As String)
        Dim i As Integer
        
        If maske = "" Then MaskT (0)
        
        If c = "c" Then
             For i = Me.OrdreView.nodes.Count To 1 Step -1
                       If Split(Me.OrdreView.nodes(i).Tag, "£")(1) = "chap" Then
                            Me.OrdreView.nodes(i).Expanded = False
                            Me.OrdreView.nodes(i).ForeColor = vbBlack
                        End If
                Next i
                If Me.OrdreView.Enabled = False Then Me.OrdreView.Enabled = True
                Me.AjPosition.Value = "Selectionner Ci dessous..."
                Me.TitrePosition.Caption = "Position "
        
        ElseIf c = "f" Then
                For i = 1 To Me.OrdreView.nodes.Count
                         If i <> Me.OrdreView.nodes.Count Then
                            If Split(Me.OrdreView.nodes(i).Tag, "£")(1) = "chap" And Split(Me.OrdreView.nodes(i + 1).Tag, "£")(1) = "fonc" Then
                                 Me.OrdreView.nodes(i).ForeColor = RGB(192, 192, 192)
                                 Me.OrdreView.nodes(i).Expanded = True
                             End If
                         Else
                         
                         End If
                  Next i
                  If Me.OrdreView.Enabled = False Then Me.OrdreView.Enabled = True
                  Me.AjPosition.Value = "Selectionner Ci dessous..."
                  Me.TitrePosition.Caption = "Position "
                  Me.OrdreView.nodes(1).EnsureVisible
            
        End If
       
       If maske = "" Then MaskT (1)
       Application.EnableEvents = True
End Function
Private Sub CommandButton1_Click()
        Me.OrdreView.nodes.Clear
        Call Load
        Me.COPIE.Caption = "Clique droit Pour Couper"
        Me.OrdreView.CheckBoxes = True
        '    Me.OrdreView.Nodes(selectedNode).Selected = False
        defautColor
        Me.chapitre.Enabled = True
        Me.fonction.Enabled = True
        Me.addItems.Visible = True
        checkList = ""
        Me.Annuler.Visible = False
        
End Sub


Private Sub OrdreView_Collapse(ByVal Node As MSComctlLib.Node)
    If Me.EtatAjout.Caption = "ADD" Then
        If Me.AjFonction.Value = True Then Node.Expanded = True
     End If
End Sub

Private Sub OrdreView_Expand(ByVal Node As MSComctlLib.Node)
    If Me.EtatAjout.Caption = "ADD" Then
        If Node.Bold = True And Me.AjChapitre.Value = True Then Node.Expanded = False
     End If
End Sub

Private Sub OrdreView_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
            If Button = 1 And Me.Pcolonne <> "Modifier" Then
                Call renitMode
            End If
End Sub

Private Sub Pcolonne_Change()
        Me.chapitre.Value = False
        Me.fonction.Value = False
      
        Me.defautColor
        If Me.Pcolonne.Value = "Modifier" Then
                Me.chapitre.Enabled = False
                Me.fonction.Enabled = False
                Me.OrdreView.CheckBoxes = False
                Me.COPIE.Caption = "Selectionner Pour Modifier"
        Else
                Me.chapitre.Enabled = True
                Me.fonction.Enabled = True
                Me.OrdreView.CheckBoxes = True
                Me.COPIE.Caption = "Clique droit Pour Couper"
        End If
End Sub

Private Sub OrdreView_Click()
      If Me.EtatAjout.Caption = "ADD" Then
       
        Exit Sub
      End If
      If Me.Pcolonne <> "Modifier" Then renitMode
End Sub

Private Sub OrdreView_NodeClick(ByVal Node As MSComctlLib.Node)
            Dim Valeur As String
            On Error GoTo Ers
            If Me.EtatAjout.Caption = "ADD" Then
                 If Me.AjChapitre.Value = False And Me.AjFonction.Value = False Then
                    Me.AjPosition.Value = "Selectionner Ci dessous..."
                    Me.TitrePosition.Caption = "Position "
                    MsgBox "Choisir Chapitre Ou Fonction", vbCritical, "BILIQUE"
                    Node.Selected = False
                    Exit Sub
                ElseIf Node.ForeColor = RGB(192, 192, 192) Then
                    MsgBox "Selectionner Fonction", vbCritical, "BILIQUE"
                    Me.AjPosition.Value = "Selectionner Ci dessous..."
                    Me.TitrePosition.Caption = "Position "
                    Exit Sub
                Else
                    Me.AjPosition.Value = Node.text
                    Me.AjPosition.Tag = Node.key
                    Me.TitrePosition.Caption = "Position : " & AvantApres(Node.text)
                    Exit Sub
                End If
            End If
            If Me.Pcolonne <> "Modifier" Then
                       If Node.index = 1 And Me.COPIE = "Selectionner Pour Coller" Then Me.OrdreView.nodes(Node.key).Selected = True
                       If Me.COPIE <> "Selectionner Pour Coller" Then Me.OrdreView.nodes(Node.key).Selected = False
                       If Me.chapitre.Value = False And Me.fonction.Value = False Then
                           MsgBox "Selectionner d'abord Group SDV", vbCritical, "BILIQUE"
                       End If
                       If eventCheck = False Then
                           Exit Sub
                       End If
                    
                       If Me.COPIE.Caption = "Selectionner Pour Coller" Then
                             If Node.ForeColor = 12632256 Then
                                   Me.OrdreView.nodes(Node.key).Selected = False
                                   MsgBox "Vous ne pouvez pas Coller Ici", vbCritical, "BILIQUE"
                             Else
                                  If CopyPaste = True Then Application.CommandBars("CopiPaste").ShowPopup
                             End If
                       End If
           Else
                Valeur = InputBox("Entrer La Valeur", "Confirmation")
                If Len(Valeur) > 0 Then
                    Call editValeur(Valeur)
                End If
           End If
           
           
Ers:
   If ERR.Number <> 0 Then
        MsgBox ERR.description, vbCritical, "BILIQUE"
    End If
End Sub
Private Sub OrdreView_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
    On Error GoTo Ers
    If Me.EtatAjout.Caption = "ADD" Then Exit Sub
     If Me.Pcolonne <> "Modifier" Then
            If Button = 2 Then
                  If Me.COPIE.Caption <> "Selectionner Pour Coller" And checkSelect(False) <> "" Then
                        If CopyPaste = True Then Application.CommandBars("CopiPaste").ShowPopup
                  ElseIf checkSelect(False) = "" Then
                                MsgBox "Cocher  D'abord Pour Faire un Choix", vbCritical, "BILIQUE"
                  End If
            End If
      End If
      
      
Ers:
   If ERR.Number <> 0 Then
        MsgBox ERR.description, vbCritical, "BILIQUE"
    End If
End Sub

Private Sub PropagateChecks(ByVal ParentNode As MSComctlLib.Node)
  Dim oNode As MSComctlLib.Node
  Dim lNodeIndex As Long
  If ParentNode.Children > 0 Then
    Set oNode = ParentNode.Child
    oNode.Checked = ParentNode.Checked
    Call PropagateChecks(oNode)
    For lNodeIndex = 1 To ParentNode.Children - 1
      Set oNode = oNode.Next
      oNode.Checked = ParentNode.Checked
      Call PropagateChecks(oNode)
    Next
  End If
  Set oNode = Nothing
End Sub
Private Sub OrdreView_NodeCheck(ByVal Node As MSComctlLib.Node)
    
    On Error GoTo Ers
     If Me.EtatAjout.Caption = "ADD" Then Exit Sub
     If Me.COPIE <> "Selectionner Pour Coller" Then Me.OrdreView.nodes(1).Selected = False
     If Me.chapitre.Value = False And Me.fonction.Value = False Then
         MsgBox "Selectionner d'abord Group SDV", vbCritical, "BILIQUE"
     End If
     If eventCheck = False Then
            
            Exit Sub
     End If
   
     PropagateChecks Node

    
Ers:
   If ERR.Number <> 0 Then
        MsgBox ERR.description, vbCritical, "BILIQUE"
    End If
End Sub



Private Sub UserForm_Activate()
       On Error GoTo Ers
       
        If checkCorrect <> "" Then
               MsgBox "Revoir Structure  doublons détèctés" & vbCrLf & checkCorrect, vbCritical, "BILIQUE"
                Unload Me
                Exit Sub
        End If
       If Me.ChargeOrder.Caption = "" Then
       
         Call Load
         Me.ChargeOrder.Caption = "OK"
         chargeD = ""
         If Me.OrdreView.nodes.Count > 0 Then Me.OrdreView.nodes(1).Selected = False
       ElseIf Me.EtatAjout = "" And chargeD <> "" Then
            MaskT (0)
            Call addItemMod(False)
            defautColor
            MaskT (1)
             If Me.OrdreView.nodes.Count > 0 Then Me.OrdreView.nodes(1).Selected = False
            chargeD = ""
       Else
                If saveSelNd <> "" Then
                
                    On Error Resume Next
                    Me.OrdreView.nodes(saveSelNd).Selected = True
                    Me.OrdreView.nodes(saveSelNd).EnsureVisible
                    Me.OrdreView.nodes(saveSelNd).Selected = False
                    ERR.Clear
                    
                End If
       End If
Ers:
   If ERR.Number <> 0 Then
        MsgBox ERR.description, vbCritical, "BILIQUE"
    End If
    
End Sub
Private Sub Load()
    Dim v As Variant
    
   
        v = defaultOrder
        loadNew (v)
    
    
    Me.Pcolonne.AddItem "Reajuster"
    Me.Pcolonne.AddItem "Modifier"
    Me.Pcolonne.Value = "Reajuster"
  
End Sub
Function loadNew(v As Variant)
    Dim lastRow As Long, i As Long
    Dim getOrder() As String
    Dim chap As String, Schap As String, fonc As String, keys As String
  
    chap = ""
    Schap = ""
    fonc = ""
    keys = ""
    
    For i = 0 To UBound(v, 1)
      If Len(v(i, 0)) > 0 Then
                If InStr(1, v(i, 1), ".") = 0 Then
                    chap = CStr(v(i, 0))
                    Me.OrdreView.nodes.Add key:=chap, text:=chap
                    Me.OrdreView.nodes(chap).Tag = v(i, 2) & "£chap"
                    Me.OrdreView.nodes(chap).Bold = True
                    Me.OrdreView.nodes(chap).Expanded = True
                    Schap = ""
                    fonc = ""
                Else
                    getOrder = Split(v(i, 1), ".")
                    If IsNumeric(getOrder(UBound(getOrder))) = False Then
                        If chap <> "" Then
                                 Schap = CStr(v(i, 0))
                                 Me.OrdreView.nodes.Add Me.OrdreView.nodes(chap).key, tvwChild, Me.OrdreView.nodes(chap).key & "|" & Schap, Schap
                                 keys = Me.OrdreView.nodes(chap).key & "|" & Schap
                                 Me.OrdreView.nodes(keys).Expanded = True
                                 Me.OrdreView.nodes(keys).Tag = v(i, 2) & "£fonc"

                         End If
                    Else
                          If UBound(Split(v(i, 1), ".")) = 1 And Schap <> "" Then Schap = ""
                          If chap <> "" And Schap <> "" Then
                               fonc = CStr(v(i, 0))
                               keys = CStr(Me.OrdreView.nodes(chap).key & "|" & Schap)
                               Me.OrdreView.nodes.Add keys, tvwChild, keys & "|" & fonc, fonc
                               keys = CStr(keys & "|" & fonc)
                               Me.OrdreView.nodes(keys).Expanded = True
                               Me.OrdreView.nodes(keys).Tag = v(i, 2) & "£fonc"
                            
                            ElseIf chap <> "" Then
                               fonc = CStr(v(i, 0))
                               Me.OrdreView.nodes.Add Me.OrdreView.nodes(chap).key, tvwChild, Me.OrdreView.nodes(chap).key & "|" & fonc, fonc
                               keys = CStr(Me.OrdreView.nodes(chap).key & "|" & fonc)
                               Me.OrdreView.nodes(keys).Expanded = True
                               Me.OrdreView.nodes(keys).Tag = v(i, 2) & "£fonc"
                             
                            End If
                    End If
                    
                End If
                 
           End If
    Next i


End Function
Function loadRead(v As Variant)
    Dim i As Integer
    For i = 0 To UBound(v, 1)
      If Len(v(i, 0)) > 0 Then
                If Len(v(i, 1)) = 0 Then
                    Me.OrdreView.nodes.Add key:=v(i, 0), text:=v(i, 0)
                    Me.OrdreView.nodes(v(i, 0)).Tag = v(i, 2) & "£chap"
                    Me.OrdreView.nodes(v(i, 0)).Bold = True
                    Me.OrdreView.nodes(v(i, 0)).Expanded = True
                
                Else
                    Me.OrdreView.nodes.Add v(i, 0), tvwChild, v(i, 0) & "|" & v(i, 1), v(i, 1)
                    Me.OrdreView.nodes(v(i, 0) & "|" & v(i, 1)).Tag = v(i, 2) & "£fonc"
                    Me.OrdreView.nodes(v(i, 0) & "|" & v(i, 1)).Expanded = True
                End If
                
           End If
    Next i


End Function

Function AvantApres(ch As String) As String
    Dim i As Integer
    Dim Lasts As String
    AvantApres = ""
    If Me.AjChapitre.Value = True Then
           For i = 1 To Me.OrdreView.nodes.Count
                        If Me.OrdreView.nodes(i).Bold = True Then
                            Lasts = Me.OrdreView.nodes(i).text
                        End If
          Next i
    ElseIf Me.AjFonction = True Then
           If Me.OrdreView.nodes(Me.AjPosition.Tag).index <> Me.OrdreView.nodes.Count Then
                 If (Split(Me.OrdreView.nodes(Me.OrdreView.nodes(Me.AjPosition.Tag).index + 1).Tag, "£")(1) = "chap") Then
                            AvantApres = "Après"
                            Exit Function
                 ElseIf (InStr(1, Me.OrdreView.nodes(Me.OrdreView.nodes(Me.AjPosition.Tag).index + 1).key, "|") <> 0 And _
                    InStr(1, Me.OrdreView.nodes(Me.OrdreView.nodes(Me.AjPosition.Tag).index).key, "|")) <> 0 Then
                       If UBound(Split(Me.OrdreView.nodes(Me.OrdreView.nodes(Me.AjPosition.Tag).index + 1).key, "|")) <> _
                          UBound(Split(Me.OrdreView.nodes(Me.OrdreView.nodes(Me.AjPosition.Tag).index).key, "|")) Then
                            AvantApres = "Après"
                            Exit Function
                       End If
                 End If
                 
           Else
                Lasts = Me.OrdreView.nodes(Me.OrdreView.nodes.Count).text
           End If
    End If
    If Lasts = ch Then AvantApres = "Après" Else AvantApres = "Avant"
End Function
Function eventCheck() As Boolean
      Dim selectedNs As String
      Dim tabCheck() As String
      Dim i As Integer
      selectedNs = checkSelect(False)
      
      eventCheck = True
      If selectedNs = "" Then Exit Function
      
      If InStr(1, selectedNs, "#") = 0 Then
            ReDim tabCheck(0)
            tabCheck(0) = selectedNs
      Else
            tabCheck = Split(selectedNs, "#")
      End If
        
      uncheck = ""
    
      
      If UBound(tabCheck) = 0 And Me.OrdreView.nodes(Split(tabCheck(0), "@")(0)).index = Me.OrdreView.nodes.Count And _
             Me.OrdreView.nodes(Split(tabCheck(0), "@")(0)).Bold = True Then
              If uncheck = "" Then uncheck = Me.OrdreView.nodes(Split(tabCheck(0), "@")(0)).key Else uncheck = uncheck & "#" & Me.OrdreView.nodes(Split(tabCheck(0), "@")(0)).key
      Else
              For i = 0 To UBound(tabCheck)
                    If Me.chapitre.Value = False And Me.fonction.Value = False Then
                         If uncheck = "" Then uncheck = Me.OrdreView.nodes(Split(tabCheck(i), "@")(0)).key Else uncheck = uncheck & "#" & Me.OrdreView.nodes(Split(tabCheck(i), "@")(0)).key
        
                    ElseIf Me.fonction.Value = True And checkRestriction(Me.OrdreView.nodes(Split(tabCheck(i), "@")(0)), "fonction") <> "" Then
                         If uncheck = "" Then uncheck = Me.OrdreView.nodes(Split(tabCheck(i), "@")(0)).key Else uncheck = uncheck & "#" & Me.OrdreView.nodes(Split(tabCheck(i), "@")(0)).key
            
                    ElseIf checkRestriction(Me.OrdreView.nodes(Split(tabCheck(i), "@")(0)), "chapitre") <> "" And Me.chapitre.Value = True Then
                        If InStr(1, Me.OrdreView.nodes(Split(tabCheck(i), "@")(0)).key, "|") <> 0 Then
                                If Me.OrdreView.nodes(Split(Me.OrdreView.nodes(Split(tabCheck(i), "@")(0)).key, "|")(0)).Checked = False Then
                                    If uncheck = "" Then uncheck = Me.OrdreView.nodes(Split(tabCheck(i), "@")(0)).key Else uncheck = uncheck & "#" & Me.OrdreView.nodes(Split(tabCheck(i), "@")(0)).key
                                End If
                        Else
                                If uncheck = "" Then uncheck = Me.OrdreView.nodes(Split(tabCheck(i), "@")(0)).key Else uncheck = uncheck & "#" & Me.OrdreView.nodes(Split(tabCheck(i), "@")(0)).key
                        End If
             
                    End If
                    
              Next i
      End If
           
      If uncheck <> "" Then eventCheck = False
End Function
Function checkRestriction(ByVal Node As MSComctlLib.Node, structure As String) As String
      checkRestriction = ""
      If structure = "chapitre" And Not Node.Parent Is Nothing Then
            checkRestriction = "Cocher Chapitre Seulement"
            Exit Function
     ElseIf structure = "fonction" And Split(Node.Tag, "£")(1) <> "fonc" Then
            checkRestriction = "Cocher Fonction Seulement"
            Exit Function
    End If

End Function

Function selectedNode() As Integer
        Dim i As Integer
        selectedNode = 0
        For i = 1 To Me.OrdreView.nodes.Count
             If Me.OrdreView.nodes(i).Selected = True Then
                selectedNode = Me.OrdreView.nodes(i).index
                Exit Function
             End If
        Next
End Function
Function checkSelect(color As Boolean) As String
        Dim i As Integer
        checkSelect = ""
        For i = 1 To Me.OrdreView.nodes.Count
             If Me.OrdreView.nodes(i).Checked = True Then
                If checkSelect = "" Then checkSelect = Me.OrdreView.nodes(i).key & "@" & Me.OrdreView.nodes(i).Tag Else checkSelect = checkSelect & "#" & Me.OrdreView.nodes(i).key & "@" & Me.OrdreView.nodes(i).Tag
                If color = True Then Me.OrdreView.nodes(i).ForeColor = vbRed
             End If
        Next
        checkList = checkSelect
End Function
Function getSelectedNode() As String
        Dim i As Integer
        getSelectedNode = ""
        For i = 1 To Me.OrdreView.nodes.Count
             If Me.OrdreView.nodes(i).ForeColor = vbRed Then
                If getSelectedNode = "" Then getSelectedNode = Me.OrdreView.nodes(i).key & "@" & Me.OrdreView.nodes(i).Tag Else getSelectedNode = getSelectedNode & "#" & Me.OrdreView.nodes(i).key & "@" & Me.OrdreView.nodes(i).Tag
             End If
        Next
        
End Function
Function defautColor()
        Dim i As Integer
       
        For i = 1 To Me.OrdreView.nodes.Count
                If Me.OrdreView.nodes(i).Bold = True Or Split(Me.OrdreView.nodes(i).Tag, "£")(1) = "chap" Then Me.OrdreView.nodes(i).Expanded = True
                Me.OrdreView.nodes(i).ForeColor = vbBlack
        Next
        If Me.OrdreView.nodes.Count > 0 Then Me.OrdreView.nodes(1).EnsureVisible
End Function
Function GreyColor()
          Dim i As Integer
          
          For i = 1 To Me.OrdreView.nodes.Count
              If Me.fonction.Value = True Then
                    If Me.OrdreView.nodes(i).Parent Is Nothing Then
                           Me.OrdreView.nodes(i).ForeColor = RGB(192, 192, 192)
                    End If
             End If
          Next i
          
          If Me.OrdreView.nodes(Me.OrdreView.nodes.Count).Bold = True Then
                Dim fonc
                Dim keys
                fonc = "Coller Ici"
              
                Me.OrdreView.nodes.Add Me.OrdreView.nodes(Me.OrdreView.nodes.Count).key, tvwChild, Me.OrdreView.nodes(Me.OrdreView.nodes.Count).key & "|" & fonc, fonc
                keys = CStr(Me.OrdreView.nodes(Me.OrdreView.nodes.Count - 1).key & "|" & fonc)
                Me.OrdreView.nodes(keys).Expanded = True
                Me.OrdreView.nodes(keys).Tag = fonc & "£fonc"
          End If
        
  
End Function
Function unckeckALL()
        Dim i As Integer
       
        For i = 1 To Me.OrdreView.nodes.Count
             If Me.OrdreView.nodes(i).Checked = True Then
                Me.OrdreView.nodes(i).Checked = False
             End If
        Next
      
End Function
Function CopyPaste() As Boolean
        Dim selectVal As String
        Dim Caption As String
        Dim targetNode As String
        
        targetNode = ""
        CopyPaste = True
        If Me.COPIE <> "Selectionner Pour Coller" Then
            Caption = "Couper"
        Else
            If checkList = "" Then checkList = getSelectedNode
            If checkPaste(False) = False Then
                MsgBox "Vous ne pouvez pas Coller Ici", vbCritical, "BILIQUE"
                CopyPaste = False
                Exit Function
            ElseIf val(selectedNode) = 0 Then
                CopyPaste = False
                Exit Function
            ElseIf checkList <> "" And Me.COPIE = "Selectionner Pour Coller" Then
                If selectedNode <> 0 Then
                    If Me.fonction = True Then
                          If Len(Me.OrdreView.nodes(val(selectedNode)).text) > 0 Then
                                If Me.OrdreView.nodes(val(selectedNode)).text = "Coller Ici" Then
                                    targetNode = Me.OrdreView.nodes(val(selectedNode)).text
                                Else
                                    targetNode = "Coller Avant " & Me.OrdreView.nodes(val(selectedNode)).text
                                End If
                          End If
                          
                    ElseIf Me.chapitre = True Then
                             If InStr(1, Me.OrdreView.nodes(val(selectedNode)).key, "|") <> 0 Then
                                targetNode = Split(Me.OrdreView.nodes(selectedNode).key, "|")(0)
                             Else
                                targetNode = Me.OrdreView.nodes(selectedNode).key
                             End If
                            
                             If checkPaste(True) = False And getChapterPosition(targetNode) = 2 Then
                                     targetNode = "Coller Aprés " & targetNode
                             Else
                                     targetNode = "Coller Avant " & targetNode
                             End If
                    End If
                    If Len(targetNode) = 0 Then
                         CopyPaste = False
                         Exit Function
                    End If
                    Caption = targetNode
                 End If
            End If
        End If

    On Error Resume Next
    Application.CommandBars("CopiPaste").Delete
    On Error GoTo 0
    Application.CommandBars.Add Name:="CopiPaste", position:=msoBarPopup, Temporary:=True
    
    With Application.CommandBars("CopiPaste").Controls.Add(msoControlButton)
        .Caption = Caption
         If Caption = "Couper" Then
            .FaceId = 3002
         Else
            .FaceId = 2985
         End If
        .OnAction = "couperColler"
    End With
    
    If Caption <> "Couper" Then
        If Me.fonction = True Then
'             With Application.CommandBars("CopiPaste").Controls.Add(msoControlButton)
'                targetNode = replace(targetNode, "Coller Avant ", "")
'                targetNode = replace(targetNode, "Coller Aprés ", "")
'                .Caption = "Coller Dans " & targetNode
'                .FaceId = 2986
'                .OnAction = "newClef"
'            End With
         End If
         
        With Application.CommandBars("CopiPaste").Controls.Add(msoControlButton)
            .Caption = "Annuler"
            .FaceId = 2983
            .OnAction = "annulerCouper"
        End With
    Else
    
'        With Application.CommandBars("CopiPaste").Controls.Add(msoControlButton)
'            .Caption = "Supprimer"
'            .FaceId = 1019
'            .OnAction = "delItems"
'        End With
        
    End If
     
End Function

Function structureChecked() As String
        structureChecked = ""
        If Me.chapitre.Value = True Then structureChecked = "Chapitre"
        If Me.fonction.Value = True Then structureChecked = "Fonction"
        
End Function
Function linkNodes()
      Dim selectedNs As String
      Dim tabCheck() As String
      Dim i As Integer
      selectedNs = checkSelect(False)
    
      If selectedNs = "" Then Exit Function
      If InStr(1, selectedNs, "#") = 0 Then
            ReDim tabCheck(0)
            tabCheck(0) = selectedNs
      Else
            tabCheck = Split(selectedNs, "#")
      End If
        
     
      For i = 0 To UBound(tabCheck)
                PropagateChecks Me.OrdreView.nodes(Split(tabCheck(i), "@")(0))
      Next i
      
    
End Function
Function renitMode()
    Dim tabCheck() As String
    Dim i As Integer
    Dim selectedNs As String
       
    On Error GoTo Ers
    
    Call eventCheck
     If uncheck <> "" Then
                 If InStr(1, uncheck, "#") = 0 Then
                        ReDim tabCheck(0)
                        tabCheck(0) = uncheck
                  Else
                        tabCheck = Split(uncheck, "#")
                  End If
                  For i = 0 To UBound(tabCheck)
                      Me.OrdreView.nodes(tabCheck(i)).Checked = False
                  Next i
        End If
    Call linkNodes
    
    selectedNs = checkSelect(False)
    If InStr(1, selectedNs, "#") = 0 Then
          ReDim tabCheck(0)
          tabCheck(0) = selectedNs
    Else
          tabCheck = Split(selectedNs, "#")
    End If
    
    If UBound(tabCheck) + 1 >= val(Me.Total) Then
           If structureChecked <> "" And uncheck <> "" Then MsgBox "Cocher " & structureChecked & " Seulement", vbCritical, "BILIQUE"
     End If
     uncheck = ""
     If selectedNs = "" Then
           Me.Total = 0
     Else
          Me.Total = UBound(tabCheck) + 1
    End If
    
Ers:
   If ERR.Number <> 0 Then
        MsgBox ERR.description, vbCritical, "BILIQUE"
    End If
     
     
End Function



Private Sub UserForm_Initialize()
    Me.EtatAjout.Caption = ""
    
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'      If Button = 1 Then
'            Call renitMode
'      End If
End Sub


Function checkPaste(Cible As Boolean) As Boolean
        Dim i As Integer
        Dim tabCheck() As String, checkedV As String
        checkedV = checkList
        checkPaste = True
        If InStr(1, checkedV, "#") = 0 Then
            ReDim tabCheck(0)
            tabCheck(0) = checkedV
        Else
            tabCheck = Split(checkedV, "#")
        End If
        
        With OrdreS.OrdreView
            For i = 0 To UBound(tabCheck)
                   If Cible = False Then
                            If Me.OrdreView.SelectedItem.key = CStr(Split(tabCheck(i), "@")(0)) Then
                                 checkPaste = False
                                 Exit Function
                            End If
                    Else
                            
                            If Me.OrdreView.nodes(CStr(Split(tabCheck(i), "@")(0))).Parent Is Nothing And Me.OrdreView.nodes(CStr(Split(tabCheck(i), "@")(0))).index = 1 Then
                                 checkPaste = False
                                 Exit Function
                            End If
                    End If
            Next i
        End With
End Function

Function getChapterPosition(Chapter As String, Optional Liste As String)
        Dim i As Integer
        Dim tot As Integer
        Dim tabCheck() As String, checkedV As String
        
       tot = 0
       getChapterPosition = 0
       If Liste = "" Then
            For i = 1 To Me.OrdreView.nodes.Count
                If Me.OrdreView.nodes(i).Parent Is Nothing Then
                    tot = tot + 1
                    If Me.OrdreView.nodes(i).key = Chapter Then
                       getChapterPosition = tot
                       Exit Function
                    End If
                End If
            Next
       Else
                checkedV = Liste
                If InStr(1, checkedV, "#") = 0 Then
                    ReDim tabCheck(0)
                    tabCheck(0) = checkedV
                Else
                    tabCheck = Split(checkedV, "#")
                End If
                         
                For i = 0 To UBound(tabCheck)
                      If InStr(1, Split(tabCheck(i), "@")(0), "|") = 0 Then
                           tot = tot + 1
                            If Split(tabCheck(i), "@")(0) = Chapter Then
                                getChapterPosition = tot
                                Exit Function
                            End If
                      End If

                Next i
       End If
End Function

Function getAllNodes(All As Boolean, Optional checkItem As String)
      Dim i As Integer
      
      getAllNodes = ""
        For i = 1 To Me.OrdreView.nodes.Count
           
                    If All = True Then
                        If getAllNodes = "" Then getAllNodes = Me.OrdreView.nodes(i).key & "@" & Me.OrdreView.nodes(i).Tag Else getAllNodes = getAllNodes & "#" & Me.OrdreView.nodes(i).key & "@" & Me.OrdreView.nodes(i).Tag
                    Else
                        If checkItem = "OK" Then
                            If Me.OrdreView.nodes(i).Checked = False Then
                                    If getAllNodes = "" Then getAllNodes = Me.OrdreView.nodes(i).key & "@" & Me.OrdreView.nodes(i).Tag Else getAllNodes = getAllNodes & "#" & Me.OrdreView.nodes(i).key & "@" & Me.OrdreView.nodes(i).Tag
                             End If
                        Else
                            If Me.OrdreView.nodes(i).ForeColor <> vbRed Then
                                    If getAllNodes = "" Then getAllNodes = Me.OrdreView.nodes(i).key & "@" & Me.OrdreView.nodes(i).Tag Else getAllNodes = getAllNodes & "#" & Me.OrdreView.nodes(i).key & "@" & Me.OrdreView.nodes(i).Tag
                             End If
                        End If
                    End If
          
        Next i
        
End Function

Function nodeExist(nodes As String) As Boolean
      Dim i As Integer
      nodeExist = False
        For i = 1 To Me.OrdreView.nodes.Count
            If UCase(Me.OrdreView.nodes(i).text) = UCase(nodes) Then
                nodeExist = True
                Exit Function
            End If
        Next i
        
End Function


Function moveCheckedValue(Optional newC As String)
    Dim i As Integer, before As Boolean
    Dim tabCheck() As String, checkedV As String, v() As String, targetNode As String, allNode As String
    Dim selNode As Integer
    On Error GoTo Ers
    
    If newC <> "DEL" Then
        If newC = "AddChapitre" Or newC = "AddFonction" Then
            selNode = Me.OrdreView.nodes(Me.AjPosition.Tag).index
        Else
            selNode = selectedNode
            If val(selNode) = 0 Then
                    MsgBox "Merci De Selectionner A Nouveau", vbCritical, "BILIQUE"
                    Exit Function
            End If
         End If
   End If

    If newC <> "DEL" Then
        If newC = "AddChapitre" Or newC = "AddFonction" Then
            checkedV = getAllNodes(True)
            allNode = checkedV
        Else
            checkedV = getAllNodes(False)
            allNode = getAllNodes(True)
        End If
    Else
        checkedV = getAllNodes(False, "OK")
    End If
    
    before = False
    If InStr(1, checkedV, "#") = 0 Then
        ReDim tabCheck(0)
        tabCheck(0) = checkedV
    Else
        tabCheck = Split(checkedV, "#")
    End If
  
    
    If newC <> "DEL" Then
                 If Me.chapitre.Value = True Then
                        If InStr(1, Me.OrdreView.nodes(val(selNode)).key, "|") <> 0 Then
                           targetNode = Split(Me.OrdreView.nodes(selNode).key, "|")(0)
                        Else
                           targetNode = Me.OrdreView.nodes(selNode).key
                        End If
                        If checkPaste(True) = False And getChapterPosition(targetNode, allNode) = 2 Then before = True
                 Else
                        targetNode = Me.OrdreView.nodes(val(selNode)).key
                 End If
                        
                Me.OrdreView.nodes.Clear
                 Me.hide
                If Me.chapitre.Value = True Or Me.AjChapitre.Value = True Then
                           For i = 0 To UBound(tabCheck)
                                
                                If newC = "AddChapitre" Then
                                                If Me.TitrePosition.Caption = "Position : Avant" Then
                                                        If getChapterPosition(CStr(Split(tabCheck(i), "@")(0)), allNode) = getChapterPosition(targetNode, allNode) Then
                                                            Call insertCheckedNode(, newC)
                                                            Call addNodeByText(CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)))
                                                        Else
                                                            Call addNodeByText(CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)))
                                                        End If
                                                ElseIf Me.TitrePosition.Caption = "Position : Après" Then
                                                        If i = UBound(tabCheck) Then
                                                            Call addNodeByText(CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)))
                                                            Call insertCheckedNode(, newC)
                                                        Else
                                                            Call addNodeByText(CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)))
                                                        End If
                                              End If
                                       
                                 Else
                                        If before = True Then
                                              If getChapterPosition(CStr(Split(tabCheck(i), "@")(0)), checkedV) = 2 Then
                                                   insertCheckedNode
                                              End If
                                              Call addNodeByText(CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)))
                                        Else
                                              If getChapterPosition(CStr(Split(tabCheck(i), "@")(0)), allNode) = getChapterPosition(targetNode, allNode) Then
                                                  insertCheckedNode
                                              End If
                                             Call addNodeByText(CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)))
                                        End If
                                    
                                End If
                         Next i
            Else
                    For i = 0 To UBound(tabCheck)
                               
                                    If Split(tabCheck(i), "@")(0) = targetNode Then
                                         If newC = "OK" Then
                                            Call addNodeByText(CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)))
                                            Call insertCheckedNode(targetNode, newC)
                                         ElseIf newC = "AddFonction" Then
                                                    If Me.TitrePosition.Caption = "Position : Avant" Then
                                                            Call insertCheckedNode(targetNode, newC)
                                                            Call addNodeByText(CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)))
                                                    ElseIf Me.TitrePosition.Caption = "Position : Après" Then
                                                            If Split(Split(tabCheck(i), "@")(1), "£")(1) = "chap" Then
                                                                Call addNodeByText(CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)))
                                                                Call insertCheckedNode(targetNode, newC)
                                                            Else
                                                                Call addNodeByText(CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)))
                                                                Call insertCheckedNode(targetNode, newC)
                                                            End If
                                                    End If
                                        Else
                                            Call insertCheckedNode(targetNode, newC)
                                            Call addNodeByText(CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)))
                                         End If
                                    Else
                                         Call addNodeByText(CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)))
                                    End If
                                    
                      Next i
             
           End If
    Else
          Me.OrdreView.nodes.Clear
          For i = 0 To UBound(tabCheck)
             Call addNodeByText(CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)))
          Next i
    End If
    
     If Me.OrdreView.nodes(Me.OrdreView.nodes.Count).text = "Coller Ici" Then
         Me.OrdreView.nodes.Remove Me.OrdreView.nodes.Count
    End If
    
     If newC <> "AddChapitre" And newC <> "AddFonction" Then
        Me.COPIE.Caption = "Clique droit Pour Couper"
        Me.OrdreView.CheckBoxes = True
        defautColor
        Me.chapitre.Enabled = True
        Me.fonction.Enabled = True
        Me.addItems.Visible = True
        checkList = ""
        Me.Annuler.Visible = False
        Application.ScreenUpdating = True
    Else
            Me.AjNom.Value = ""
            Me.AjPosition.Value = "Selectionner Ci dessous..."
            If newC = "AddChapitre" Then Call chargeChapFonct("c", "x")
            If newC = "AddFonction" Then Call chargeChapFonct("f", "x")
    End If
    saveSelNd = targetNode
    On Error Resume Next
    Me.Show
    
    ERR.Clear
    On Error GoTo Ers
    
Ers:
   If ERR.Number <> 0 Then
        MsgBox ERR.description, vbCritical, "BILIQUE"
        If Me.OrdreView.Visible = False Then MaskT (1)
    End If
End Function
Function insertCheckedNode(Optional entete As String, Optional newClefs As String)
        Dim i As Integer
        Dim tabCheck() As String, checkedV As String
        Dim replaceKeys As String, repl As String, jointTable As String
        
        If newClefs = "AddChapitre" Or newClefs = "AddFonction" Then
            If Me.AjFonction.Value = True Then checkedV = Me.AjNom.Value & "@" & "Nouveau" & replace(Time, ":", "") & "£fonc"
            If Me.AjChapitre.Value = True Then checkedV = Me.AjNom.Value & "@" & "Nouveau" & replace(Time, ":", "") & "£chap"
       Else
            checkedV = checkList
       End If
        
        If InStr(1, checkedV, "#") = 0 Then
            ReDim tabCheck(0)
            tabCheck(0) = checkedV
        Else
            tabCheck = Split(checkedV, "#")
        End If
                
        jointTable = Join(tabCheck, ";")
        If entete <> "" Then
            If newClefs <> "OK" Then
                    If newClefs = "AddFonction" And InStr(1, entete, "|") = 0 Then
                        entete = entete
                    Else
                        entete = Left(entete, InStrRev(entete, "|") - 1)
                    End If
            End If
            
        End If
    
        
        With OrdreS.OrdreView
            For i = 0 To UBound(tabCheck)
                       If entete <> "" Then
                            If newClefs = "OK" Then
                                Call insertNodeByText(entete, CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)))
                            ElseIf newClefs = "AddFonction" Then
                                If InStr(1, CStr(Split(tabCheck(i), "@")(0)), "|") = 0 Then
                                     Call insertNodeByText(entete, "Nouveau|" & CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)), jointTable)
                                Else
                                    Call insertNodeByText(entete, CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)), jointTable)
                                End If
                            Else
                                Call insertNodeByText(entete, CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)), jointTable)
                            End If
'                            Call addNodeByText(CStr(entete & "|" & Right(Split(tabCheck(i), "@")(0), Len(Split(tabCheck(i), "@")(0)) - InStrRev(Split(tabCheck(i), "@")(0), "|"))), CStr(Split(tabCheck(i), "@")(1)))
                       Else
                            Call addNodeByText(CStr(Split(tabCheck(i), "@")(0)), CStr(Split(tabCheck(i), "@")(1)))
                       End If
'                      .Nodes.Add .Nodes(After).key, tvwChild, Right(tabCheck(i), Len(tabCheck(i)) - InStr(1, tabCheck(i), "|")), Right(tabCheck(i), Len(tabCheck(i)) - InStr(1, tabCheck(i), "|"))
            Next i
        End With
        
        
End Function

Function addNodeByText(textAdd As String, tags As String, Optional textKey As String)
        Dim chem As String
        Dim Schap As String
        Dim checkExistc As String
        
         If InStr(1, textAdd, "|") <> 0 Then
                If Split(tags, "£")(1) = "fonc" Then
                    Me.OrdreView.nodes.Add Left(textAdd, InStrRev(textAdd, "|") - 1), tvwChild, textAdd, Split(textAdd, "|")(UBound(Split(textAdd, "|")))
                    Me.OrdreView.nodes(textAdd).Expanded = True
                    Me.OrdreView.nodes(textAdd).Tag = tags
                End If
         Else
                Me.OrdreView.nodes.Add key:=textAdd, text:=textAdd
                Me.OrdreView.nodes(textAdd).Bold = True
                Me.OrdreView.nodes(textAdd).Expanded = True
                Me.OrdreView.nodes(textAdd).Tag = tags
        End If
        
End Function
Function insertNodeByText(textKey As String, textAdd As String, tags As String, Optional joinTable As String)
        Dim chem As String
        Dim Schap As String
        Dim checkExistc As String
        Dim tabV() As String
        Dim i As Integer
        
        Schap = Left(textAdd, InStrRev(textAdd, "|") - 1)
        chem = replace(textAdd, Schap & "|", "")
        Schap = ""
        tabV = Split(textAdd, "|")
        For i = 0 To UBound(tabV) - 1
             
             If joinTable <> "" Then
                 If getCheckedItem(joinTable, tabV(i)) = True Then
                     Schap = IIf(Schap = "", tabV(i), Schap & "|" & tabV(i))
                     On Error Resume Next
                     checkExistc = Me.OrdreView.nodes(textKey & "|" & Schap).Tag
                 End If
             Else
                   Schap = IIf(Schap = "", tabV(i), Schap & "|" & tabV(i))
                   On Error Resume Next
                   checkExistc = Me.OrdreView.nodes(textKey & "|" & Schap).Tag
             End If
             
             If ERR.Number <> 0 Then
                ERR.Clear
                If InStr(1, Schap, "|") = 0 Then
                    Schap = replace(Schap, tabV(i), "")
                Else
                    Schap = replace(Schap, "|" & tabV(i), "")
                End If
                
             End If
             
        Next i
        
        If ERR.Number <> 0 Then ERR.Clear
        If Schap <> "" Then
            Me.OrdreView.nodes.Add (textKey & "|" & Schap), tvwChild, textKey & "|" & Schap & "|" & chem, chem
            Me.OrdreView.nodes(textKey & "|" & Schap & "|" & chem).Expanded = True
            Me.OrdreView.nodes(textKey & "|" & Schap & "|" & chem).Tag = tags
        Else
            Schap = Right(textAdd, Len(textAdd) - InStrRev(textAdd, "|"))
            Me.OrdreView.nodes.Add (textKey), tvwChild, textKey & "|" & Schap, Schap
            Me.OrdreView.nodes(textKey & "|" & Schap).Expanded = True
            Me.OrdreView.nodes(textKey & "|" & Schap).Tag = tags
        End If
                
                            
End Function
Function getCheckedItem(joinTable As String, itemCompare As String) As Boolean
        Dim i As Integer
        Dim tabV() As String
        
        getCheckedItem = False
        If InStr(1, joinTable, ";") = 0 Then
            ReDim tabV(0)
            tabV(0) = joinTable
        Else
            tabV = Split(joinTable, ";")
        End If
        
    
        For i = 0 To UBound(tabV)
            If Split(Right(tabV(i), Len(tabV(i)) - InStrRev(tabV(i), "|")), "@")(0) = itemCompare Then
                 getCheckedItem = True
                 Exit Function
            End If
        Next i
        
End Function


Function editValeur(v As String)
        Dim getC As String, getIndex As Integer
        
        getC = Me.OrdreView.SelectedItem.key
        getIndex = Me.OrdreView.nodes(getC).index
        If nodeExist(v) = False Then
                If InStr(1, getC, "|") <> 0 Then
                        If UBound(Split(Me.OrdreView.nodes(getIndex).key, "|")) = 2 Then
                             Me.OrdreView.nodes(getIndex).key = Left(getC, InStrRev(getC, "|") - 1) & "|" & v
                             Me.OrdreView.SelectedItem.text = v
                        Else
                             Me.OrdreView.nodes(getIndex).key = Split(getC, "|")(0) & "|" & v
                             Me.OrdreView.SelectedItem.text = v
                             Call replaceKey(CStr(Split(getC, "|")(1)), v, 2)
                        End If
                    Else
                         Me.OrdreView.nodes(getIndex).key = v
                         Me.OrdreView.SelectedItem.text = v
                         Call replaceKey(getC, v, 1)
                    End If
        Else
            MsgBox "Valeur existe déjà", vbCritical, "BILIQUE"
        End If
End Function

Function replaceKey(search As String, replace As String, order As Integer)
       Dim i As Integer
    
        For i = 1 To Me.OrdreView.nodes.Count
              If InStr(1, Me.OrdreView.nodes(i).key, "|") <> 0 Then
                        If UBound(Split(Me.OrdreView.nodes(i).key, "|")) = 2 Then
                            If order = 2 Then
                                    If Split(Me.OrdreView.nodes(i).key, "|")(1) = search Then
                                        Me.OrdreView.nodes(i).key = _
                                        Split(Me.OrdreView.nodes(i).key, "|")(0) & "|" & replace & "|" & Split(Me.OrdreView.nodes(i).key, "|")(2)
                                    End If
                              ElseIf order = 1 Then
                                     If Split(Me.OrdreView.nodes(i).key, "|")(0) = search Then
                                        Me.OrdreView.nodes(i).key = _
                                        replace & "|" & Split(Me.OrdreView.nodes(i).key, "|")(1) & "|" & Split(Me.OrdreView.nodes(i).key, "|")(2)
                                    End If
                              End If
                       ElseIf UBound(Split(Me.OrdreView.nodes(i).key, "|")) = 1 Then
                              
                              If order = 1 Then
                                     If Split(Me.OrdreView.nodes(i).key, "|")(0) = search Then
                                        Me.OrdreView.nodes(i).key = _
                                        replace & "|" & Split(Me.OrdreView.nodes(i).key, "|")(1)
                                    End If
                              End If
                              
                       End If
            End If
        Next i
End Function

Function validateAll()
        Dim i As Long, totAdd As Integer
        Dim getField As Object, getSaut As Object, colon As Object
        Dim j As Long
        Dim colVals(7) As String
        Dim derniereLigne As Long
        Dim derniereColonne As Long, cm As Integer
        Dim RsRows As Object
        Dim c
        Dim SE As String
        Dim checkSaut As String
        Dim saveCol As Object
        checkSaut = ""
        totAdd = 0
        Set getSaut = CreateObject("Scripting.Dictionary")
        Set cOrder = CreateObject("Scripting.Dictionary")
        createOrder
        Dim rrW As Integer
        
        With ThisWorkbook.sheets("SDV MANAGER")
          
                            Set getField = CreateObject("Scripting.Dictionary")
                            
                            For i = 1 To Me.OrdreView.nodes.Count
                                  If Not getField.Exists(CStr(Split(Me.OrdreView.nodes(i).Tag, "£")(0))) Then
                                             getField.Add key:=CStr(Split(Me.OrdreView.nodes(i).Tag, "£")(0)), Item:=Me.OrdreView.nodes(i).text & "@" & i
                                  End If
                                  
                            Next i
                    
'                             Set RsRows = CreateObject("ADODB.Recordset")
'                             RsRows.ActiveConnection = db.GetOdb
'                             RsRows.Properties("Jet OLEDB:Locking Granularity") = 1
'                             RsRows.Open "[" & .Range("B3") & "]", db.GetOdb, 1, 3, 2
'
'                             Set colon = CreateObject("Scripting.Dictionary")
'                             For Each c In RsRows.fields
'                                       If Not colon.Exists(UCase(c.Name)) Then
'                                              colon.Add Key:=UCase(c.Name), Item:=UCase(c.Name)
'                                          End If
'                              Next c
                            
                            derniereColonne = 2
                            For cm = 1 To derniereColonne
                               If .Cells(.Rows.Count, cm).End(xlUp).row > derniereLigne Then derniereLigne = .Cells(.Rows.Count, cm).End(xlUp).row
                            Next cm
                            If derniereLigne <= 4 Then
                                    MsgBox "inf"
                                    Exit Function
                            End If
                          
                            .Range("A2:B" & derniereLigne).ClearContents
                            .Range("A2:B" & derniereLigne).Interior.color = RGB(255, 255, 255)
                            .Range("A2:B" & derniereLigne).Font.color = RGB(0, 0, 0)
                             
                             For i = 1 To Me.OrdreView.nodes.Count
                                      
                                                   
                                                    .Range("B" & i + 1) = cOrder(CStr(Split(Me.OrdreView.nodes(i).Tag, "£")(0)))
                                                    .Range("A" & i + 1) = Split(getField(CStr(Split(Me.OrdreView.nodes(i).Tag, "£")(0))), "@")(0)
                                                    If InStr(1, .Range("B" & i + 1), ".") = 0 Then
                                                        .Range("A" & i + 1).Interior.color = 11851260
                                                    End If
                              Next i
                                                        
                          MAJRatingPosition
                       
                          
          End With
End Function
Function returnRow(RsRows As Object, Tag As Variant, ord As Integer, TagF As Variant, Optional colA As Object) As Variant
                    Dim colon As Object
                    Dim i As Integer
                    Dim fFind As Boolean
                    RsRows.movefirst
                    Set colon = CreateObject("Scripting.Dictionary")
                    i = 0
                    
                  
                    If colA Is Nothing Then
                            While Not RsRows.EOF
                                  
                                  If Not IsNull(RsRows("Aide Générale").Value) Then
                                          If Not colon.Exists(RsRows("field_order").Value) Then
                                                 colon.Add key:=RsRows("field_order").Value, Item:=i
                                           End If
                                           i = i + 1
                                  End If
                                  If RsRows("field_order").Value = Tag Then
                                       RsRows("order") = ord
                                       RsRows("field_order") = TagF
                                       RsRows.Update
'                                       Exit Function
                                   ElseIf IsNull(RsRows("Aide Générale").Value) Then
                                           RsRows.Delete
                                           
                                   End If
                                 
                                 RsRows.MoveNext
                            Wend
                            
                            Set returnRow = colon
                   Else
                        If Len(colA(Tag)) > 0 Then
                            RsRows.Move CLng(val(colA(Tag)))
                            RsRows("order") = ord
                            RsRows("field_order") = TagF
                            RsRows.Update
                       End If
                   End If
End Function
Function updatekf(RsRows As Object, Optional colA As Object) As Variant
                    Dim i As Integer
                    RsRows.movefirst
                   
                   For i = 1 To Me.OrdreView.nodes.Count
                  
                           If Left(CStr(Split(Me.OrdreView.nodes(i).Tag, "£")(0)), 7) = "Nouveau" Then
                                RsRows.addnew
                               
                                RsRows("field_order") = cOrder(CStr(Split(Me.OrdreView.nodes(i).Tag, "£")(0)))
                                RsRows("Aide Générale") = Me.OrdreView.nodes(i).text
                                
                                RsRows("order") = val(colA(CStr(Split(Me.OrdreView.nodes(i).Tag, "£")(0))))
                                If CStr(Split(Me.OrdreView.nodes(i).Tag, "£")(1)) = "chap" Then
                                    RsRows("font_color") = 16777215
                                    RsRows("bg_color") = 8210719
                                Else
                                    RsRows("font_color") = 0
                                    RsRows("bg_color") = 16777215
                                End If
                                RsRows.Update
                            End If
                  Next i
                                
End Function
Function delKF(RsRows As Object, Optional gf As Object)
      Dim i As Integer
     RsRows.movefirst
     
     
     If Not gf Is Nothing Then
             While Not RsRows.EOF
                   If Not gf.Exists(RsRows("field_order").Value) And IsNull(RsRows("field_order").Value) = False Then
                      
                        RsRows.Delete
                        
                   End If
                   RsRows.MoveNext
            Wend
    End If
End Function
Function addHEader(RsRows As Object, derniereColonne As Long, colon As Object)
                     Dim i As Integer
                     Dim j As Integer
                     Dim colVals(7)
                     
                     colVals(1) = "Aide Générale"
                    colVals(2) = "Version 7"
                    colVals(3) = "Surveillance"
                    colVals(4) = "Statistique"
                    colVals(5) = "Label"
                    colVals(6) = "Indice Zone"
                    colVals(7) = "Statistique multizones"
                     With ThisWorkbook.sheets("SYNTHESE")
                     For i = 2 To 3
                                 If Application.CountA(.Rows(i).EntireRow) > 0 Then
                                         RsRows.addnew
                                          For j = 2 To derniereColonne
                                                 If colon.Exists(replace(UCase(colVals(j - 1)), ".", "|")) Then
                                                                 RsRows(replace((colVals(j - 1)), ".", "|")) = .Cells(i, j - 1).Value
                                                  End If
                                         Next j
                                        
                                         RsRows("font_color") = .Cells(i, 1).Font.color
                                         RsRows("bg_color") = .Cells(i, 1).Interior.color
                                          RsRows("order") = i
                                         RsRows.Update
                                                    
                                                        
                                  End If
                       Next i
         End With
                   

End Function

Function createOrder()
    Dim i As Long
    Dim saves As Integer
    Dim savesLoop As Integer
    Dim chapOrder(3) As Integer
    Dim strConcatOrder As String
    Dim saveConcatOrder As Object
    Dim getSC As String
    Dim j As Integer
    Set saveConcatOrder = CreateObject("Scripting.Dictionary")
    chapOrder(1) = 0
    chapOrder(2) = 0
    chapOrder(3) = 0
    

     For i = 1 To Me.OrdreView.nodes.Count
              If Me.OrdreView.nodes(i).Bold = True Then
                            chapOrder(1) = chapOrder(1) + 1
                            chapOrder(2) = 0
                            chapOrder(3) = 0
                            saves = 0
                            If Not cOrder.Exists(Split(Me.OrdreView.nodes(i).Tag, "£")(0)) Then
                                     cOrder.Add key:=Split(Me.OrdreView.nodes(i).Tag, "£")(0), Item:=intToOrder(chapOrder(1), True)
                            End If
                            If saveConcatOrder.Count > 0 Then saveConcatOrder.RemoveAll
                            strConcatOrder = intToOrder(chapOrder(1), True)
                ElseIf Split(Me.OrdreView.nodes(i).Tag, "£")(1) = "fonc" Then
                
                            savesLoop = UBound(Split(Me.OrdreView.nodes(i).key, "|"))
                            getSC = Left(Me.OrdreView.nodes(i).key, InStrRev(Me.OrdreView.nodes(i).key, "|") - 1)

                            If saveConcatOrder.Exists(savesLoop) Then
                                If saveConcatOrder.Exists(savesLoop & ";" & getSC) Then
                                    chapOrder(3) = Split(saveConcatOrder(savesLoop & ";" & getSC), ".")(UBound(Split(saveConcatOrder(savesLoop & ";" & getSC), ".")))
                                    strConcatOrder = saveConcatOrder(savesLoop & ";" & getSC)
                                    strConcatOrder = Left(strConcatOrder, InStrRev(strConcatOrder, ".") - 1)
                                Else
                                    chapOrder(3) = 0
                                     If savesLoop > 1 Then
                                            strConcatOrder = saveConcatOrder(savesLoop - 1)
                                    Else
                                           strConcatOrder = saveConcatOrder(savesLoop)
                                           strConcatOrder = Left(strConcatOrder, InStrRev(strConcatOrder, ".") - 1)
                                    End If
                                    
                                End If
                            Else
                                chapOrder(3) = 0
                            End If
                            j = i
                            While i <= Me.OrdreView.nodes.Count And savesLoop = UBound(Split(Me.OrdreView.nodes(j).key, "|"))
                                        chapOrder(3) = chapOrder(3) + 1
                                        If Not cOrder.Exists(Split(Me.OrdreView.nodes(i).Tag, "£")(0)) Then
                                                 cOrder.Add key:=Split(Me.OrdreView.nodes(i).Tag, "£")(0), Item:=strConcatOrder & "." & chapOrder(3)
                                        End If
                                        i = i + 1
                                        If i <= Me.OrdreView.nodes.Count Then j = i
                            Wend
                            strConcatOrder = strConcatOrder & "." & chapOrder(3)
                            i = i - 1
                            
                            If Not saveConcatOrder.Exists(savesLoop & ";" & getSC) Then
                                    saveConcatOrder.Add key:=savesLoop & ";" & getSC, Item:=strConcatOrder
                             Else
                                    saveConcatOrder.Remove key:=savesLoop & ";" & getSC
                                    saveConcatOrder.Add key:=savesLoop & ";" & getSC, Item:=strConcatOrder
                             End If
                             
                            
                                    
                             If Not saveConcatOrder.Exists(savesLoop) Then
                                    saveConcatOrder.Add key:=savesLoop, Item:=strConcatOrder
                              Else
                                    saveConcatOrder.Remove key:=savesLoop
                                    saveConcatOrder.Add key:=savesLoop, Item:=strConcatOrder
                             End If
                    
              End If
                 
    Next i
    
End Function

Function makeBold()
        Dim i As Integer
      
        For i = 1 To Me.OrdreView.nodes.Count
            If Split(Me.OrdreView.nodes(i).Tag, "£")(1) = "chap" Then
                Me.OrdreView.nodes(i).Bold = True
            End If
        Next i
        
End Function







