Attribute VB_Name = "PopPup"
Option Explicit

Function couperColler()
    If OrdreS.COPIE.Caption <> "Selectionner Pour Coller" Then
         OrdreS.COPIE.Caption = "Selectionner Pour Coller"
         OrdreS.checkSelect (True)
         OrdreS.OrdreView.CheckBoxes = False
         OrdreS.chapitre.Enabled = False
         OrdreS.fonction.Enabled = False
         OrdreS.addItems.Visible = False
         OrdreS.GreyColor
         OrdreS.Annuler.Visible = True
    Else
        Call OrdreS.moveCheckedValue
    End If
End Function
Function newClef()
    Call OrdreS.moveCheckedValue("OK")
End Function
Function delItems()
    If MsgBox("Voulez Vous Supprimer ?", vbCritical + vbYesNo, "BILIQUE") = vbYes Then
         Call OrdreS.moveCheckedValue("DEL")
    End If
End Function
Function annulerCouper()
    OrdreS.COPIE.Caption = "Clique droit Pour Couper"
    OrdreS.OrdreView.CheckBoxes = True
    OrdreS.defautColor
    OrdreS.chapitre.Enabled = True
    OrdreS.fonction.Enabled = True
    OrdreS.addItems.Visible = True
    OrdreS.Annuler.Visible = False
     If OrdreS.OrdreView.nodes(OrdreS.OrdreView.nodes.Count).text = "Coller Ici" Then
         OrdreS.OrdreView.nodes.Remove OrdreS.OrdreView.nodes.Count
    End If
End Function

Sub hh()
 Dim v As Variant
    Dim lastRow As Long, i As Long
    Dim getOrder() As String
    Dim chap As String, Schap As String, fonc As String, keys As String
    chap = ""
    Schap = ""
    fonc = ""
    keys = ""

    
    v = defaultOrder
    For i = 0 To UBound(v, 1)
      If Len(v(i, 1)) > 0 Then
                If InStr(1, v(i, 1), ".") = 0 Then
                    chap = CStr(v(i, 0))
                    OrdreS.OrdreView.nodes.Add key:=chap, text:=chap
                    OrdreS.OrdreView.nodes(chap).Tag = v(i, 2)
                    OrdreS.OrdreView.nodes(chap).Bold = True
                    OrdreS.OrdreView.nodes(chap).Expanded = True
                    Schap = ""
                    fonc = ""
                Else
                    getOrder = Split(v(i, 1), ".")
                    If IsNumeric(getOrder(UBound(getOrder))) = False Then
                        If chap <> "" Then
                                 Schap = CStr(v(i, 0))
                                 OrdreS.OrdreView.nodes.Add OrdreS.OrdreView.nodes(chap).key, tvwChild, OrdreS.OrdreView.nodes(chap).key & "|" & Schap, Schap
                                 keys = OrdreS.OrdreView.nodes(chap).key & "|" & Schap
                                 OrdreS.OrdreView.nodes(keys).Expanded = True
                                 OrdreS.OrdreView.nodes(keys).Tag = v(i, 2)
                         End If
                    Else
                          If chap <> "" And Schap <> "" Then
                               fonc = CStr(v(i, 0))
                               keys = CStr(OrdreS.OrdreView.nodes(chap).key & "|" & Schap)
                               OrdreS.OrdreView.nodes.Add keys, tvwChild, keys & "|" & fonc, fonc
                               keys = CStr(keys & "|" & fonc)
                               OrdreS.OrdreView.nodes(keys).Expanded = True
                               OrdreS.OrdreView.nodes(keys).Tag = v(i, 2)
                            ElseIf chap <> "" Then
                               fonc = CStr(v(i, 0))
                               OrdreS.OrdreView.nodes.Add OrdreS.OrdreView.nodes(chap).key, tvwChild, OrdreS.OrdreView.nodes(chap).key & "|" & fonc, fonc
                               keys = CStr(OrdreS.OrdreView.nodes(chap).key & "|" & fonc)
                               OrdreS.OrdreView.nodes(keys).Expanded = True
                               OrdreS.OrdreView.nodes(keys).Tag = v(i, 2)
                            End If
                    End If
                    
                End If
                 
           End If
    Next i
        

    OrdreS.Pcolonne.AddItem "Reajuster"
    OrdreS.Pcolonne.AddItem "Modifier"
    OrdreS.Pcolonne.Value = "Reajuster"
  
End Sub




