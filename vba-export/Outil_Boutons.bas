Attribute VB_Name = "Outil_Boutons"
' Cellule Office

Option Explicit

Sub ViewPop()
    
    If Not ValidateMainDatabase() Then
        MsgBox "DriveScope database not found." & vbCrLf & GetMainDbPath(), vbCritical, "DriveScope"
        Exit Sub
    End If
    
    'If ThisWorkbook.Sheets("HOME").Range("Targets").Value = "" Or ThisWorkbook.Sheets("HOME").Range("Milestone").Value = "" Or ThisWorkbook.Sheets("HOME").Range("Gears") = "" Or ThisWorkbook.Sheets("HOME").Range("Fuel").Value = "" Then
    '    MsgBox "Information missing. Please complete the ""PROJECT SUMMARY"" section.", vbExclamation, "Warning!"
    'Else
    form.Show
    db.Terminate
    'End If
End Sub

Sub Update_Button()
    Dim msg_err
    
    On Error GoTo Ers
    EventAndScreen (False)
    ProgressLoad
    ProgressTitle ("MAJ Targets")
    
   
    Set colorGlobalDriv = CreateObject("Scripting.Dictionary")
    Set colorGlobalDrivPred = CreateObject("Scripting.Dictionary")
    
    Call affect_WTP(ActiveSheet.Name, "driv")
    
    If checkCriteria(ThisWorkbook.ActiveSheet.Name) = True And checkCorrespondancePriority(ThisWorkbook.ActiveSheet.Name) = True Then
            Call UpdateTab(ThisWorkbook.ActiveSheet.Name, "driv")
            Call NoteGlobale_2
            Call Moniteur("""" & ThisWorkbook.ActiveSheet.Name & """ tab has been updated")
    End If
    Call MaskEmptySdv
    Range("B1").Select
    ProgressExit
    EventAndScreen (True)
    
    Set colorGlobalDriv = Nothing
    Set colorGlobalDrivPred = Nothing
   
    
Ers:
    If ERR.Number <> 0 Then
        msg_err = ERR.description
        ProgressExit
        EventAndScreen (True)
        MsgBox msg_err, vbCritical, "ODRIV"
    End If
End Sub
Sub Update_ButtonDyn()
    Dim msg_err
    
    On Error GoTo Ers
    EventAndScreen (False)
    ProgressLoad
    ProgressTitle ("MAJ Targets")
    
    Set colorGlobalDyn = CreateObject("Scripting.Dictionary")
    Set colorGlobalDynPred = CreateObject("Scripting.Dictionary")
  
    
    Call affect_WTP(ActiveSheet.Name, "dyn")

    If checkCriteriaDyn(ThisWorkbook.ActiveSheet.Name) = True And checkCorrespondancePriorityDyn(ThisWorkbook.ActiveSheet.Name) = True Then
            Call UpdateTab(ThisWorkbook.ActiveSheet.Name, "dyn")
            Call NoteGlobale_2
            Call Moniteur("""" & ThisWorkbook.ActiveSheet.Name & """ tab has been updated")
    End If
     Call MaskEmptySdv
    Range("Bi1").Select
    ProgressExit
    EventAndScreen (True)
    
    Set colorGlobalDyn = Nothing
    Set colorGlobalDynPred = Nothing
  
    
Ers:
    If ERR.Number <> 0 Then
        msg_err = ERR.description
        ProgressExit
        EventAndScreen (True)
        MsgBox msg_err, vbCritical, "ODRIV"
    End If
End Sub

Sub UpdateTarget_Button()
    Application.ScreenUpdating = False
     Call affect_WTP(ActiveSheet.Name, "driv")
     Call Couleur_Points(ActiveSheet.Name, SDV2Ncrit(ActiveSheet.Name))
     HideC3
    Range("B1").Select
    Application.ScreenUpdating = True
End Sub

Sub UpdateTarget_ButtonDyn()
    Application.ScreenUpdating = False
     Call affect_WTP(ActiveSheet.Name, "dyn")
     Call Couleur_PointsDyn(ActiveSheet.Name, SDV2Ncrit(ActiveSheet.Name))
     HideC3
    Range("BI1").Select
    Application.ScreenUpdating = True
End Sub

Sub Exit_Button()
    Dim current As String

    Application.ScreenUpdating = False
    current = ThisWorkbook.ActiveSheet.Name
    ThisWorkbook.sheets("HOME").Select
    If current = "VERSIONS" Then
        ThisWorkbook.sheets(current).Visible = False
    End If
    Application.ScreenUpdating = True
End Sub

Sub YellowGreenOff_Button()
    Application.ScreenUpdating = False
    Call DisplayRedOnly(ThisWorkbook.ActiveSheet.Name)
    Application.ScreenUpdating = True
End Sub
Sub YellowGreenOff_ButtonDyn()
    Application.ScreenUpdating = False
    Call DisplayRedOnlyDyn(ThisWorkbook.ActiveSheet.Name)
    Application.ScreenUpdating = True
End Sub

Sub GreenOff_Button()
    Application.ScreenUpdating = False
    Call DisplayRedYellow(ThisWorkbook.ActiveSheet.Name)
    Application.ScreenUpdating = True
End Sub
Sub GreenOff_ButtonDyn()
    Application.ScreenUpdating = False
    Call DisplayRedYellowDyn(ThisWorkbook.ActiveSheet.Name)
    Application.ScreenUpdating = True
End Sub
Sub FilterPopul_Button()
    Call PopulFilter(ThisWorkbook.ActiveSheet.Name)
End Sub
Sub FilterPopul_ButtonDYN()
    Call PopulFilterDyn(ThisWorkbook.ActiveSheet.Name)
End Sub
Sub UnlockSheetS()
        unlocksheet.Show
End Sub
Sub ChangeTargets_Button()

    If ThisWorkbook.sheets("HOME").Range("Software").Value = "" Or ThisWorkbook.sheets("HOME").Range("Milestone").Value = "" Or ThisWorkbook.sheets("HOME").Range("Gears") = "" Or ThisWorkbook.sheets("HOME").Range("Fuel").Value = "" Then
        MsgBox "Information missing. No project exist.", vbExclamation, "Warning!"
    Else
        'bouton "SETTINGS" pour modifier les informations
        'Unload ProjectInfo
        
        Load ProjectInfo
        ProjectInfo.Show
        db.Terminate
    End If
End Sub

Sub GenerateReport_Button()
    If ThisWorkbook.sheets("HOME").Range("Software").Value = "" Or ThisWorkbook.sheets("HOME").Range("Milestone").Value = "" Or ThisWorkbook.sheets("HOME").Range("Gears") = "" Or ThisWorkbook.sheets("HOME").Range("Fuel").Value = "" Then
        MsgBox "Information missing. Please complete the ""PROJECT SUMMARY"" section.", vbExclamation, "Warning!"
    Else
        ThisWorkbook.sheets("DocVersions").Range("D7:D10").ClearContents
        ThisWorkbook.sheets("DocVersions").Visible = True
        ThisWorkbook.sheets("DocVersions").Activate

        ThisWorkbook.sheets("DocVersions").Range("D7").Select

        DoEvents
    End If
End Sub

Sub StartNewProject_Button()
   
    If ThisWorkbook.sheets("structure").Range("N1").Value > 2 Or IsError(Evaluate("='SUBJECTIVE'!A1")) = False Or ThisWorkbook.sheets("HOME").Range("Project").Value <> "" Then
        
        If MsgBox("Are you sure? This action will close the current project.", vbYesNo + vbCritical, "Create New Project") = vbYes Then
            Call Erase_All2
            MsgBox "Project information. Informations of project successfully registered.", vbInformation, "ODRIV"
             
            Unload frmNameCode
            Unload New_Project
            Load New_Project
            New_Project.Show
        End If
    Else
        Unload frmNameCode
        Unload New_Project
        Load New_Project
        New_Project.Show
             
    End If
    db.Terminate
End Sub


Sub AddFile_Button()
    On Error GoTo gEr
    Dim wb As Workbook
    Dim msg_err
    
    If Not ValidateMainDatabase() Then
        MsgBox "DriveScope database not found." & vbCrLf & GetMainDbPath(), vbCritical, "DriveScope"
        Exit Sub
    End If
        
    Set wb = ActiveWorkbook
    EventAndScreen (False)
    
    ThisWorkbook.sheets("HOME").Range("Prestation").Value = "DRIVABILITY"
    If ThisWorkbook.sheets("HOME").Range("Fuel").Value = "" Or ThisWorkbook.sheets("HOME").Range("Gears").Value = "" Or ThisWorkbook.sheets("HOME").Range("Prestation").Value = "" Or ThisWorkbook.sheets("HOME").Range("Software").Value = "" Or ThisWorkbook.sheets("HOME").Range("Milestone").Value = "" Or ThisWorkbook.sheets("HOME").Range("Area").Value = "" Then
        MsgBox "Project information missing. Please Start new project before adding files.", vbCritical
    Else
        Call LoadData
       
    End If
    sheets("HOME").Activate
    EventAndScreen (True)
    db.Terminate
gEr:
    If ERR.Number <> 0 Then
        msg_err = ERR.description
        EventAndScreen (True)
        Application.DisplayAlerts = False
        If sheetExists("DATA") Then ThisWorkbook.sheets("DATA").Delete
        Application.DisplayAlerts = True
        ProgressExit
        wb.sheets("HOME").Activate
        MsgBox msg_err, vbCritical, "ODRIV"
        db.Terminate
    End If
End Sub

Sub AddSubjective_Button()
   
End Sub

Sub Instructions_Button()
    Application.ScreenUpdating = False
    ThisWorkbook.sheets("DNT").OLEObjects("Instructions").Verb
    ThisWorkbook.sheets("HOME").Activate
    Application.ScreenUpdating = True

    Call Moniteur("Instructions have been open")
End Sub

Sub Versions_Button()
    ThisWorkbook.sheets("VERSIONS").Visible = True
    ThisWorkbook.sheets("VERSIONS").Activate
    Call Moniteur("""VERSIONS"" tab has been displayed")
End Sub

Sub StatsTITO_Button()
   
End Sub

Sub XY_TITO_Button()
'    Call XY_TITO(ThisWorkbook.ActiveSheet.name)
End Sub

Sub HideC3(Optional SDV As String, Optional typeOff As String)
    Dim i As Integer

    If Len(SDV) > 2 Then
         If Len(typeOff) > 0 Then
            Call showC3(SDV, typeOff)
         Else
            Call showC3(SDV)
         End If
'         ThisWorkbook.Sheets(SDV).Shapes("FILTERS").Top = 90
'          ThisWorkbook.Sheets(SDV).Shapes("FILTERS").Left = 1140
'          ThisWorkbook.Sheets(SDV).Shapes("C_SCALE").Top = 90
'          ThisWorkbook.Sheets(SDV).Shapes("C_SCALE").Left = 1210
        If typeOff = "driv" Then
             For i = 13 To (ThisWorkbook.sheets(SDV).Range("A6:BA6").Find("Indice", , , xlPart).Column)
                If ThisWorkbook.sheets(SDV).Cells(5, i).Value = 3 Then
                    ThisWorkbook.sheets(SDV).Columns(i).EntireColumn.Hidden = True
                End If
            Next i
         ElseIf typeOff = "dyn" Then
             For i = 72 To (ThisWorkbook.sheets(SDV).Range("BH6:GG6").Find("Indice", , , xlPart).Column)
                If ThisWorkbook.sheets(SDV).Cells(5, i).Value = 3 Then
                    ThisWorkbook.sheets(SDV).Columns(i).EntireColumn.Hidden = True
                End If
            Next i
         Else
            For i = 13 To (ThisWorkbook.sheets(SDV).Range("A6:BA6").Find("Indice", , , xlPart).Column)
                If ThisWorkbook.sheets(SDV).Cells(5, i).Value = 3 Then
                    ThisWorkbook.sheets(SDV).Columns(i).EntireColumn.Hidden = True
                End If
            Next i
            
             For i = 72 To (ThisWorkbook.sheets(SDV).Range("BH6:GG6").Find("Indice", , , xlPart).Column)
                If ThisWorkbook.sheets(SDV).Cells(5, i).Value = 3 Then
                    ThisWorkbook.sheets(SDV).Columns(i).EntireColumn.Hidden = True
                End If
            Next i
            
         End If
    Else
          Call showC3
'         ActiveSheet.Shapes("FILTERS").Top = 90
'          ActiveSheet.Shapes("FILTERS").Left = 1140
'          ActiveSheet.Shapes("C_SCALE").Top = 90
'          ActiveSheet.Shapes("C_SCALE").Left = 1210
        If typeOff = "driv" Then
             For i = 13 To (ThisWorkbook.ActiveSheet.Range("A6:BA6").Find("Indice", , , xlPart).Column)
                If ThisWorkbook.ActiveSheet.Cells(5, i).Value = 3 Then
                    ThisWorkbook.ActiveSheet.Columns(i).EntireColumn.Hidden = True
                End If
            Next i
        ElseIf typeOff = "dyn" Then
             For i = 72 To (ThisWorkbook.ActiveSheet.Range("BH6:GG6").Find("Indice", , , xlPart).Column)
                If ThisWorkbook.ActiveSheet.Cells(5, i).Value = 3 Then
                    ThisWorkbook.ActiveSheet.Columns(i).EntireColumn.Hidden = True
                End If
            Next i
        Else
             If ActiveSheet.Range("A2:BG2").EntireColumn.Hidden = False Then
                     For i = 13 To (ThisWorkbook.ActiveSheet.Range("A6:BA6").Find("Indice", , , xlPart).Column)
                        If ThisWorkbook.ActiveSheet.Cells(5, i).Value = 3 Then
                            ThisWorkbook.ActiveSheet.Columns(i).EntireColumn.Hidden = True
                        End If
                    Next i
            Else
                     For i = 72 To (ThisWorkbook.ActiveSheet.Range("BH6:GG6").Find("Indice", , , xlPart).Column)
                        If ThisWorkbook.ActiveSheet.Cells(5, i).Value = 3 Then
                            ThisWorkbook.ActiveSheet.Columns(i).EntireColumn.Hidden = True
                        End If
                    Next i
           End If
           
        End If
       
    End If
    ActiveWindow.ScrollColumn = 1
    
End Sub

Sub BackToRating()
    Application.EnableEvents = False
    ThisWorkbook.sheets("RATING").Activate
    Application.EnableEvents = True
End Sub
Sub showC3(Optional SDV As String, Optional typeOff As String)
    If Len(SDV) > 2 Then
        If typeOff = "driv" Then
             ThisWorkbook.sheets(SDV).Columns("M:BD").EntireColumn.Hidden = False
        ElseIf typeOff = "dyn" Then
            ThisWorkbook.sheets(SDV).Columns("BT:GG").EntireColumn.Hidden = False
        Else
            ThisWorkbook.sheets(SDV).Columns("M:BD").EntireColumn.Hidden = False
            ThisWorkbook.sheets(SDV).Columns("BT:GG").EntireColumn.Hidden = False
        End If
       
'        ThisWorkbook.Sheets(SDV).Shapes("FILTERS").Top = 11.24992
'        ThisWorkbook.Sheets(SDV).Shapes("FILTERS").Left = 1470
'        ThisWorkbook.Sheets(SDV).Shapes("C_SCALE").Top = 11.24992
'        ThisWorkbook.Sheets(SDV).Shapes("C_SCALE").Left = 1540
    Else
        If typeOff = "driv" Then
             ThisWorkbook.ActiveSheet.Columns("M:BD").EntireColumn.Hidden = False
        ElseIf typeOff = "dyn" Then
            ThisWorkbook.ActiveSheet.Columns("BT:GG").EntireColumn.Hidden = False
        Else
             If ActiveSheet.Range("A2:BG2").EntireColumn.Hidden = False Then
               ThisWorkbook.ActiveSheet.Columns("M:BD").EntireColumn.Hidden = False
            Else
               ThisWorkbook.ActiveSheet.Columns("BT:GG").EntireColumn.Hidden = False
            End If
        End If
      
'         ActiveSheet.Shapes("FILTERS").Top = 11.24992
'        ActiveSheet.Shapes("FILTERS").Left = 1470
'        ActiveSheet.Shapes("C_SCALE").Top = 11.24992
'        ActiveSheet.Shapes("C_SCALE").Left = 1540
    End If
    ActiveWindow.ScrollColumn = 1
End Sub





