Attribute VB_Name = "Update_Tab"
' Cellule Office

Option Explicit


Sub UpdateTab(ByVal onglet As String, Optional partIndex As String)
    Dim NEVENTS As Variant
    Dim Ncrit As Integer
    Dim Goo As Boolean
    Dim gea As Integer
    
    With ThisWorkbook.sheets("SETTINGS")
         If .Cells.Find(onglet, .Cells(2, 1), xlValues, xlWhole, xlByRows, , False) Is Nothing Then
                Exit Sub
        End If
    End With
        
    Ncrit = SDV2Ncrit(onglet)
    NEVENTS = ThisWorkbook.sheets("structure").Range("N1").Value
    ProgressTitle ("Nettoyage Des SDV")
    gea = 0
    
    If Len(partIndex) = 0 Then
        Call RAZ_onglet(onglet)
        Call RAZ_ongletDyn(onglet)
    ElseIf partIndex = "driv" Then
         Call RAZ_onglet(onglet)
    ElseIf partIndex = "dyn" Then
         Call RAZ_ongletDyn(onglet)
    End If
    
    With ThisWorkbook.sheets(onglet)
        .Visible = xlSheetVisible
        If NEVENTS > 2 Then
            If Len(partIndex) = 0 Then
                If checkCriteria(onglet) = True And checkCorrespondancePriority(onglet) = True Then
                    Call updateDrivability(onglet)
                End If
                If checkCriteriaDyn(onglet) = True And checkCorrespondancePriorityDyn(onglet) = True Then
                        Call updateDynamic(onglet)
                End If
            ElseIf partIndex = "driv" Then
                Call updateDrivability(onglet)
            ElseIf partIndex = "dyn" Then
                 Call updateDynamic(onglet)
            End If
        
        End If
         
        
        If Len(partIndex) = 0 Or partIndex = "driv" Then
            Call initDriv(onglet, "driv")
        ElseIf partIndex = "dyn" Then
             Call initDriv(onglet, "dyn")
        End If
        
    End With
    
End Sub

Function updateDrivability(onglet As String)
    Dim NEVENTS As Variant
    Dim Ncrit As Integer
    Dim Goo As Boolean
    Dim gea As Integer
        
    Set colorGlobalDriv = CreateObject("Scripting.Dictionary")
    Set colorGlobalDrivPred = CreateObject("Scripting.Dictionary")
    
    Ncrit = SDV2Ncrit(onglet)
    With ThisWorkbook.sheets(onglet)
                ThisWorkbook.Worksheets(onglet).UsedRange.EntireRow.Hidden = False
                Application.StatusBar = onglet & " : Remplissage_Population_Drivability"
                ThisWorkbook.sheets(onglet).Cells.EntireColumn.Hidden = False
                Call Remplissage_Population(onglet)
               
                If TotEventSheet(onglet) > 6 Then Goo = True Else Goo = False
                gea = ThisWorkbook.sheets("HOME").Range("H23").Value
                        
                If Goo = True Then
                    ProgressTitle (onglet & " : Couleur_Points_Drivability")
                    Call affect_WTP(onglet, "driv")
                    Call Couleur_Points(onglet, Ncrit)
                 
                    ProgressTitle (onglet & " : Distribution_Drivability")
                    Call Distributions(onglet, "driv")
                    
                    ProgressTitle (onglet & " : Priorisation_Drivability")
                    Call Priorisations(onglet)
                    Call calculTPB(onglet)
                   
                    ProgressTitle (onglet & " : MAJ Graphique_Drivability")
                    Call updateGraphSdv(onglet)
                    Call convertGraph.poliGraph(onglet)
                    Call GraphToValue(onglet)
                    Call orderCol(onglet)
                    
                    ProgressTitle (onglet & " : Calcul IndiceAgrement_Drivability")
                    Call IndiceAgrement(onglet, "driv")
                  
                    ProgressTitle (onglet & " : Calcul Note_SDV_Drivability")
                     Call Note_SDV(onglet, False)
                      If sheets("HOME").Range("Milestone") <> 4 Then Call Note_SDV(onglet, True)

                    ProgressTitle (onglet & " : Calcul CRITICITY_Drivability")
                    Call F_criticity(onglet, "driv")
                   
                    If ThisWorkbook.sheets("RATING").Range("E11") <> "" Then
                       ProgressTitle (onglet & " : Calcul Note Globale ")
                       Call NoteGlobale
                   End If
                   
'                   ThisWorkbook.sheets("Graph_status").Range("I2") = GetTaux
                
           End If
           Call CamFull(onglet, "driv")
        
         If Not ThisWorkbook.sheets(onglet).Rows(6).Find(What:="Acquisition Name", LookAt:=xlWhole) Is Nothing Then
             ThisWorkbook.sheets(onglet).Columns(ThisWorkbook.sheets(onglet).Rows(6).Find(What:="Acquisition Name", LookAt:=xlWhole).Column).AutoFit
        End If
    End With
  
    Application.StatusBar = False
End Function
Function updateDynamic(onglet As String)
    Dim NEVENTS As Variant
    Dim Ncrit As Integer
    Dim Goo As Boolean
    Dim gea As Integer
    
    Ncrit = SDV2Ncrit(onglet)
   
    With ThisWorkbook.sheets(onglet)
               ThisWorkbook.Worksheets(onglet).UsedRange.EntireRow.Hidden = False
               Application.StatusBar = onglet & " : Remplissage_Population_Dynamic"
               ThisWorkbook.sheets(onglet).Cells.EntireColumn.Hidden = False
                
               Call Remplissage_PopulationDyn(onglet)
               
              If TotEventSheet(onglet) > 6 Then Goo = True Else Goo = False
              gea = ThisWorkbook.sheets("HOME").Range("H23").Value
                        
             If Goo = True Then
                    ProgressTitle (onglet & " : Couleur_Points ")
                    Call affect_WTP(onglet, "dyn")
                    Call Couleur_PointsDyn(onglet, Ncrit)
                 
                    ProgressTitle (onglet & " : Distribution_Dynamic")
                    Call Distributions(onglet, "dyn")
                    
                    ProgressTitle (onglet & " : Priorisation_Dynamic")
                    Call PriorisationsDyn(onglet)
                    Call calculTPBDyn(onglet)
                    
                    ProgressTitle (onglet & " : MAJ Graphique_Dynamic")
                    Call updateGraphSdvDyn(onglet)
                    Call convertGraph.poliGraphDyn(onglet)
                    Call GraphToValueDyn(onglet)
                    Call orderColDyn(onglet)
                     
                    ProgressTitle (onglet & " : Calcul IndiceAgrement_Dynamic")
                    Call IndiceAgrement(onglet, "dyn")
                  
                    ProgressTitle (onglet & " : Calcul Note_SDV_Dynamic")
                    Call Note_SDVDyn(onglet, False)
                    If sheets("HOME").Range("Milestone") <> 4 Then Call Note_SDVDyn(onglet, True)

                    ProgressTitle (onglet & " : Calcul CRITICITY_Dynamic")
                    Call F_criticity(onglet, "dyn")
                    
                      If ThisWorkbook.sheets("RATING").Range("E17") <> "" Then
                       ProgressTitle (onglet & " : Calcul Note Globale ")
                       Call NoteGlobaleDyn
                   End If
                   
'                   ThisWorkbook.sheets("Graph_status").Range("I20") = GetTauxDyn
                   
          End If
         
         Call CamFull(onglet, "dyn")
        
         If Not ThisWorkbook.sheets(onglet).Range("BH6:GG6").Find(What:="Acquisition Name", LookAt:=xlWhole) Is Nothing Then
             ThisWorkbook.sheets(onglet).Columns(ThisWorkbook.sheets(onglet).Range("BH6:GG6").Find(What:="Acquisition Name", LookAt:=xlWhole).Column).AutoFit
        End If
    End With
   
    Application.StatusBar = False
End Function
Sub Color_RED_PLUS(ByVal onglet As String, part As String)
    Dim r As Range
    Dim j As Integer
    Dim i As Long
    Dim rang As String
    Dim iDPart As Integer
    
    If part = "driv" Then
        iDPart = 0
    Else
        iDPart = 1
    End If
    
    rang = "A6:BA6;BH6:GG6"
    
        Set r = ThisWorkbook.sheets(onglet).Cells(7, ThisWorkbook.sheets(onglet).Range(Split(rang, ";")(iDPart)).Cells.Find(What:="Event Rating", LookAt:=xlWhole).Column)
        Do While Not isEmpty(r)
           For i = 13 To (ThisWorkbook.sheets(onglet).Range(Split(rang, ";")(iDPart)).Find("Indice", , , xlPart).Column)
                If IsNumeric(ThisWorkbook.sheets(onglet).Cells(5, i).Value) And ThisWorkbook.sheets(onglet).Cells(5, i).Value <> 3 _
                And ThisWorkbook.sheets(onglet).Cells(r.row, i).Interior.color = RGB(222, 0, 0) And r.Interior.color = RGB(255, 0, 0) Then
                    r.Value = "RED +"
                End If
            Next i
            Set r = r.Offset(1, 0)
        Loop
   
End Sub
Sub Color_RED_PLUS_BY_ROW(ByVal onglet As String, rRow As Long, iDPart As Integer)
    Dim r As Range
    Dim i As Long
    Dim rang As String
    Dim j As Integer
  
    
    rang = "A6:BA6;BH6:GG6"
    
       If iDPart = 1 Then i = 13 Else i = 72
        Set r = ThisWorkbook.sheets(onglet).Cells(rRow, ThisWorkbook.sheets(onglet).Range(Split(rang, ";")(iDPart - 1)).Cells.Find(What:="Event Rating", LookAt:=xlWhole).Column)
        If Not r Is Nothing Then
           For i = i To (ThisWorkbook.sheets(onglet).Range(Split(rang, ";")(iDPart - 1)).Find("Indice", , , xlPart).Column)
                If IsNumeric(ThisWorkbook.sheets(onglet).Cells(5, i).Value) And ThisWorkbook.sheets(onglet).Cells(5, i).Value <> 3 _
                And ThisWorkbook.sheets(onglet).Cells(r.row, i).Interior.color = RGB(222, 0, 0) And r.Interior.color = RGB(255, 0, 0) Then
                    r.Value = "RED +"
                End If
            Next i
            Set r = r.Offset(1, 0)
        End If
   
End Sub

























