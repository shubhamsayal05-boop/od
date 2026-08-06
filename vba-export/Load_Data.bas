Attribute VB_Name = "Load_Data"
' Cellule Office

'Option Explicit
Private tempGearbox As Long
Private defMode As String
Private colon As Object
Private colFound As Object
Private v As Variant
Sub AjouterSiNonExisteDansTrie(ByRef wsTrie As Worksheet)

    Dim wsEntete As Worksheet

    Dim lastCol As Long, i As Long

    Dim val1 As String, val2 As String, concatVal As String

    Dim lastRowTrie As Long

    Dim trouve As Boolean

    Dim cell As Range
 
    

    Set wsEntete = ThisWorkbook.sheets("ENTETE_COLONNE")
 
    

    lastCol = wsTrie.Cells(2, wsTrie.Columns.Count).End(xlToLeft).Column
 

    For i = 1 To lastCol

        val1 = Trim(wsTrie.Cells(1, i).Value)

        val2 = Trim(wsTrie.Cells(2, i).Value)

        concatVal = val1 & ", " & val2
 

        trouve = False

        For Each cell In wsEntete.Range("A1", wsEntete.Cells(wsEntete.Rows.Count, 1).End(xlUp))

            If LCase(Trim(cell.Value)) = LCase(concatVal) Then

                trouve = True

                Exit For

            End If

        Next cell
 
        

        If Not trouve Then

            lastRowTrie = wsEntete.Cells(wsEntete.Rows.Count, 1).End(xlUp).row + 1

            wsEntete.Cells(lastRowTrie, 1).Value = concatVal

        End If

    Next i
 
    

End Sub


Sub LoadData()
    Dim NEVENTS As Double
    Dim Chosenfile As Variant
    Dim SourceFile As String, Sourcetab As String
    Dim i As Double, Totalevents As Double
    Dim evt As Double, k As Double
    Dim c As Range, delRange As Range, rng As Range
    Dim Wsh As Worksheet
    Dim ncz As String
    Dim LeverMeca As Boolean, LeverElec As Boolean
    Dim valGet As String, lastRow As Integer
    Dim part1() As String, vG As String, stGetTemp As String
    Dim Name_Tend As String, Name_Tstart As String, versionFile As Long
    Dim RqOdb As Object
    Dim idc As String
    Dim wsheet As Worksheet
   
   
    'Choix du fichier
    Chosenfile = Application.GetOpenFilename("Excel files ,*.xlsx,Excel files,*.xlsm,Excel files,*.xl,", , "Please choose the acquisition file", , False)
    
    If Chosenfile = False Then
        Exit Sub
    End If
    ProgressLoad
    ProgressTitle ("Debut D'analyse")
    If sheetExists("DATA") Then
        Application.DisplayAlerts = False
        ThisWorkbook.sheets("DATA").Delete
        If sheetExists("GRILLE") Then ThisWorkbook.sheets("GRILLE").Delete
        Application.DisplayAlerts = True
    End If
    
    sheets.Add.Name = "DATA"
    NEVENTS = ThisWorkbook.Worksheets("DATA").Range("A65000").End(xlUp).row
    
    'Afficher le message "Please Wait"
    ThisWorkbook.Worksheets("HOME").Range("Moniteur").Interior.color = RGB(255, 0, 0)

    'Ouvrir le fichier
    SourceFile = Application.Workbooks.Open(Chosenfile).Name
    Sourcetab = "TRIE"
    ProgressTitle ("Recuperation des colonnes")
    Set wsheet = Workbooks(SourceFile).Worksheets(Sourcetab)
    Call AjouterSiNonExisteDansTrie(wsheet)
    
    v = Workbooks(SourceFile).Worksheets(Sourcetab).UsedRange.Value
    Set colFound = CreateObject("Scripting.Dictionary")
    With Workbooks(SourceFile).Worksheets(Sourcetab)
               On Error Resume Next
               versionFile = .Rows(2).Find(What:="DRIVE Version", LookIn:= _
                                 xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:= _
                                 xlNext, MatchCase:=False, SearchFormat:=False).Column
                part1 = Split(.Cells(3, versionFile).Value, ".")
                    
               If ERR.Number <> 0 Then
                    ERR.Clear
                    Call importAborted(SourceFile)
                    Exit Sub
               ElseIf "V" & part1(0) & "." & part1(1) <> db.GetValue("select version from projet where ID=" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value) Then
                    MsgBox "Files must be of same version" & vbCr & "Please re-edit your ""GrilleCotation"" files to get matching versions." _
                        & vbCr & "File Version : " & "V" & part1(0) & "." & part1(1) & vbCr & "Project Version : " & db.GetValue("select version from projet where ID=" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value), vbCritical
                    Call importAborted(SourceFile)
                    Exit Sub
               Else
                       version = .Range(nomcol("product:", "DRIVE Version", SourceFile, v) & 3).Value
                       If version = "" And InStr(1, version, ".") = 0 Then
                                 MsgBox "No Drive Version", vbCritical, "ODRIV"
                                 Call importAborted(SourceFile)
                                 Exit Sub
                       End If
               End If
               On Error GoTo 0
            
              ProgressTitle ("Selection des données")
              ThisWorkbook.Worksheets("DATA").Visible = True
              Call FiltresOff(SourceFile)
                    i = 3
                    If .Range("C" & i) = "" Then
                        MsgBox "NO DATA FOUND", vbCritical, "ODRIV"
                        Call importAborted(SourceFile)
                        Exit Sub
                    Else
                        stGetTemp = getTemperatureAllow
                        If Len(stGetTemp) > 0 Then
                            ncz = nomColFound(Split(stGetTemp, ",")(0), Split(stGetTemp, ",")(1), SourceFile, v)
                        Else
                            ncz = ""
                        End If
                        tempGearbox = 0
                        Set Wsh = Workbooks(SourceFile).Worksheets(Sourcetab)
                      
                        
                        Call TemperatureB
                        ProgressTitle ("Verification des donnees")
                        Do Until .Range("C" & i).Value = ""
                              If ncz <> "" Then
                                  If Len(.Range(ncz & i)) = 0 And (colon.Exists(UCase(.Range("C" & i).Value)) Or .Range("C" & i) Like "*lever change*") Then
                                        If tempGearbox = 0 Then
                                            While tempGearbox = 0
                                                ThisWorkbook.sheets("HOME").Activate
                                                defineTemp.Show
    '                                            .Activate
                                             Wend
                                         End If
                                         .Range(ncz & i) = tempGearbox
                                  End If
                             End If
                            valGet = conditionSDV(i, Wsh, SourceFile)
                            If valGet <> "" Then
                               .Range("C" & i).Value = valGet
                            Else
                                If delRange Is Nothing Then Set delRange = .Range("C" & i) Else Set delRange = Union(delRange, .Range("C" & i))
                               
                            End If
                            i = i + 1
                        Loop
                    End If
                    ProgressTitle ("Verification des donnees")
                    If Not delRange Is Nothing Then
                         delRange.EntireRow.Delete
                    End If
                    
                    If Wsh.Range("C65000").End(xlUp).row = 2 Then
                        MsgBox "Aucune Ligne compatible avec ODRIV", vbCritical, "ODRIV"
                        Call importAborted(SourceFile)
                        Exit Sub
                    End If
                    
                    .Range("A3:CAA" & i).Sort Key1:=.Range("C3"), Key2:=.Range("E3"), Key3:=.Range("D3")
                    NEVENTS = ThisWorkbook.Worksheets("DATA").Range("A65000").End(xlUp).row
                    .UsedRange.Copy Destination:=ThisWorkbook.Worksheets("DATA").Range("A1")
                    ThisWorkbook.sheets("DATA").Copy After:=ThisWorkbook.sheets("DATA")
                    ThisWorkbook.sheets("DATA (2)").Name = "GRILLE"
                    ProgressTitle ("Chargement des données dans la base")
                    Call InsertDB.chargeVal(, nomColFound("Sous situation de vie", "Sub Event Name", SourceFile, v))
                    ThisWorkbook.Worksheets("Structure").Range("N1") = (ThisWorkbook.Worksheets("DATA").Range("A65000").End(xlUp).row) + 7
                    Call Moniteur("New project has been created with file " & Chosenfile)
                    Application.DisplayAlerts = False
                    Workbooks(SourceFile).Close False
                    Application.DisplayAlerts = True
                    
                End With
                                
                lastRow = ThisWorkbook.Worksheets("DATA").Cells(Rows.Count, "ANW").End(xlUp).row
                Call RAZ_SDV_Sheets
                Call eraseRating
                Call eraseGraphStatus
                ProgressTitle ("Creation SDV")
                ThisWorkbook.Worksheets("VIERGE").Visible = True
                With ThisWorkbook.Worksheets("structure")
                         Set c = ThisWorkbook.sheets("structure").Range("B2")
                           Do While Len(c.Offset(0, 1).Value) > 0
                                If Len(c.Value) > 0 Then
                                   If Not ThisWorkbook.sheets("DATA").Columns(3).Cells.Find(What:=c.Value, LookAt:=xlWhole) Is Nothing Then
                                     Call GenSdVSheets(c.Value)
                                   End If
                               End If
                          Set c = c.Offset(1, 0)
                        Loop
                End With
                ThisWorkbook.Worksheets("VIERGE").Visible = False
                Call FiltresOff
            
                With ThisWorkbook
                    With .Worksheets("DATA")
                        Name_Tend = nomcol("Val_A_Tend", "Selector_Lever_Position")
                        Name_Tstart = nomcol("Val_A_Tstart", "Selector_Lever_Position")
                        If NEVENTS <= 2 Then
                            NEVENTS = 2
                        End If
                        Totalevents = .Range("A65000").End(xlUp).row
'                        .Activate
                        LeverElec = False
                        LeverMeca = False
                        evt = NEVENTS
                      
                        While LeverMeca = False And LeverElec = False And evt <= Totalevents
                            evt = evt + 1
                            If (.Range(Name_Tend & evt).Value = 4 Or .Range(Name_Tstart & evt).Value = 4) And .Range(Name_Tend & evt).Value <> .Range(Name_Tstart & evt).Value Then
                                LeverMeca = True
                            ElseIf (.Range(Name_Tend & evt).Value = 0 Or .Range(Name_Tstart & evt).Value = 0) And .Range(Name_Tend & evt).Value <> .Range(Name_Tstart & evt).Value Then
                                LeverElec = True
                            End If
                        Wend
            
                        If LeverMeca = True And LeverElec = False Then
                            For k = NEVENTS + 1 To Totalevents
                                .Range(Name_Tend & k) = .Range(Name_Tend & k) - 1
                                .Range(Name_Tstart & k) = .Range(Name_Tstart & k) - 1
                            Next k
                        End If
                    End With
                    ProgressTitle ("MAJ Target")
                    .Worksheets("HOME").Range("Project") = ThisWorkbook.sheets("DATA").Range(nomcol("Vehicle Configuration", "Vehicle Configuration Name") & 3)
                    .Worksheets("DATA").UsedRange.Offset(1, 0).Interior.color = RGB(255, 255, 255)
                    .Worksheets("DATA").UsedRange.AutoFilter
                    .Worksheets("DATA").Rows(1).AutoFilter
'                    Set rng = ThisWorkbook.Worksheets("DATA").Rows("B1:CAA65000").Find(What:="PHEV Vehicle Mode")
'                    If Not rng Is Nothing Then
                    Mode = colModeConfig
'                    End If
'                    MsgBox ThisWorkbook.Worksheets("DATA").Range(nomcol("Val_A_Tend", "PHEV Vehicle Mode") & 3).Value
                    defMode = ""
                    If Len(Mode) = 0 Then
                        While Len(defMode) = 0
                            defaultMode.Show
                        Wend
                        Mode = defMode
                    End If
'                   version = ThisWorkbook.Worksheets("DATA").Range(nomcol("product:", "DRIVE Version") & 3).Value
                   part1 = Split(version, ".")
                    vG = "V" & part1(0) & "." & part1(1)
                    
                    .Worksheets("HOME").Range("Mode") = Mode
                     idc = getDbId(ThisWorkbook.Worksheets("Home").Range("idProjects"))
                     Set RqOdb = db.GetOdb(val(idc))
    
                    If db.GetValue("select mode from projet where ID=" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value) = "" Then _
                     db.Execute ("update projet set mode='" & .Worksheets("HOME").Range("Mode") & "' Where ID=" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value)
                     Call db.Execute("update projet set mode='" & .Worksheets("HOME").Range("Mode") & "' Where ID=" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value, RqOdb)
            
                    .Worksheets("HOME").Range("DriveVersion") = vG
                 
                     If db.GetValue("select version from projet where ID=" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value) = "" Then _
                     db.Execute ("update projet set Version='" & .Worksheets("HOME").Range("DriveVersion") & "' Where ID=" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value)
                     Call db.Execute("update projet set Version='" & .Worksheets("HOME").Range("DriveVersion") & "' Where ID=" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value, RqOdb)
            
                    nom = ThisWorkbook.sheets("HOME").Range("Fuel").Value & " PREMIUM " & ThisWorkbook.sheets("HOME").Range("Prestation").Value & " (" & vG & ")"
                    Call MajAll_WTP(Mode, nom)
                    
                    Call Moniteur("File """ & SourceFile & """ has been added to your project.")
    End With

    ThisWorkbook.Worksheets("HOME").Activate
    ThisWorkbook.Worksheets("DATA").Visible = xlSheetHidden
    ThisWorkbook.Worksheets("GRILLE").Visible = xlSheetHidden
    Erase v
    Set colFound = Nothing
     ProgressExit
     db.CloseSudbConn
     MsgBox "Done", vbInformation, "ODRIV"

End Sub

Function TemperatureB()
       Dim r As Range
       Dim onglet  As String
       Set colon = CreateObject("Scripting.Dictionary")
       onglet = ""
       Set r = ThisWorkbook.sheets("structure").Cells(2, 3)
        While Not Len(r.Value) = 0
            If r.Value = "sheets" Then onglet = r.Offset(0, -1).Value
            If r.Offset(0, 1).Value = "Gearbox Temperature" Then
                    If Not colon.Exists(UCase(onglet)) Then
                        colon.Add key:=UCase(onglet), Item:=UCase(onglet)
                        If correspondance(onglet) <> "" Then If Not colon.Exists(UCase(correspondance(onglet))) Then colon.Add key:=UCase(UCase(correspondance(onglet))), Item:=UCase(correspondance(onglet))
                        If onglet = "Lever change" Then
                            If Not colon.Exists(UCase(correspondance("Engage"))) Then colon.Add key:=UCase("Engage"), Item:=UCase(correspondance("Engage"))
                            If Not colon.Exists(UCase(correspondance("Disengage"))) Then colon.Add key:=UCase("Disengage"), Item:=UCase(correspondance("Disengage"))
                        End If
                    End If
            End If
            Set r = r.Offset(1, 0)
        Wend
       
End Function
Function getTemperatureAllow()
    Dim c As Range
    
     With ThisWorkbook.Worksheets("CONFIGURATIONS")
         Set c = .Range("GEARBOX")
         Set c = c.Offset(1, 0)
         getTemperatureAllow = ""
         While c.Value <> ""
             If UCase(c.Value) = UCase(ThisWorkbook.Worksheets("HOME").Range("Gears").Value) And UCase(c.Offset(0, 4)) = "X" Then
                 If InStr(1, c.Offset(0, 5), ",") <> 0 Then
                        getTemperatureAllow = c.Offset(0, 5)
                 End If
                 Exit Function
             End If
             Set c = c.Offset(1, 0)
         Wend
    End With
       
End Function

Function searchMode(sVal As String)
    Dim c As Range
    
     With ThisWorkbook.Worksheets("CONFIGURATIONS")
         Set c = .Range("MODESCONFIG")
         Set c = c.Offset(1, 0)
         searchMode = ""
         While c.Value <> ""
             If UCase(c.Value) = UCase(sVal) Or UCase(c.Offset(0, 1).Value) = UCase(sVal) Then
                 searchMode = c.Offset(0, 2)
                 Exit Function
             End If
             Set c = c.Offset(1, 0)
         Wend
    End With
End Function
Function colModeConfig()
    
    Dim c As Range
    Dim getMode As String
     With ThisWorkbook.Worksheets("CONFIGURATIONS")
         Set c = .Range("COLMODESCONFIG")
         Set c = c.Offset(1, 0)
         colModeConfig = ""
         While c.Value <> ""
            getMode = searchMode(ThisWorkbook.Worksheets("DATA").Range(nomcol(c.Value, c.Offset(0, 1)) & 3))
            If getMode <> "" Then
                 colModeConfig = getMode
                 Exit Function
             End If
             Set c = c.Offset(1, 0)
         Wend
    End With
End Function
Function MkTemp(Temp As Long)
      tempGearbox = Temp
End Function

Function MKdefMode(Temp As String)
      defMode = Temp
End Function


Function correspondance(val As String)
                correspondance = ""
               If val = "Power-on downshift" Then
                   correspondance = "Low torque Power-on downshift"
                ElseIf val = "Sailing Exit" Then
                    correspondance = "Exit sailing"
                ElseIf val = "Drive Away Standing Start" Then
                   correspondance = "Standing start"
                ElseIf val = "Drive Away Creep" Then
                    correspondance = "Creep"
                ElseIf val = "Cold Coast - brake-on downshift" Then
                    correspondance = "Coast / brake-on downshift"
                 ElseIf val = "Coast - brake-on downshift" Then
                   correspondance = "Coast / brake-on downshift"
                 ElseIf val = "Cold Power-on upshift" Then
                    correspondance = "Power-on upshift"
                 ElseIf val = "Coast - brake-on upshift" Then
                    correspondance = "Coast / brake-on upshift"
                ElseIf val = "(PT) KD - tip in downshift" Then
                    correspondance = "Kick down / tip in downshift"
                ElseIf val = "(TO) KD - tip in downshift" Then
                   correspondance = "Kick down / tip in downshift"
                End If
                  
End Function

Function conditionSDV(ligne As Double, Wsh As Worksheet, SourceFile As Variant) As String
      Dim Ongls As String
      Dim j As Long
      Dim cnd As Boolean
      Dim TabSV() As String
      Dim t As Integer
      Dim valGet As String
      Dim p As Integer
      Dim cGet
      Dim findV As Boolean
      
       cGet = ThisWorkbook.sheets("DEFINITION SDV").UsedRange.Columns("A:E").Value
      cnd = True
      conditionSDV = ""
      
      With ThisWorkbook.Worksheets("DEFINITION SDV")
             findV = False
              For j = 1 To UBound(cGet, 1)
            
                    If IsNumeric(cGet(j, 1)) = True And Len(cGet(j, 1)) > 0 And Len(cGet(j, 3)) = 0 And .Cells(j + 2, 2).Interior.color = RGB(255, 255, 255) Then
                        Ongls = cGet(j, 2)
                        
                        j = j + 2
                         While cnd = True And j <= UBound(cGet, 1)
                                
                                If cGet(j, 3) = "CONTIENT" Then
                                         If findV = True And cGet(j, 5) = cGet(j - 1, 5) Then
                                            j = j + 1
                                       Else
                                            findV = False
                                             If InStr(1, cGet(j, 4), ";") <> 0 Then
                                                 TabSV = Split(cGet(j, 4), ";")
                                             Else
                                                 ReDim TabSV(0)
                                                 TabSV(0) = cGet(j, 4)
                                             End If
                                             valGet = ""
                                             p = IIf(j + 1 > UBound(cGet, 1), UBound(cGet), j + 1)
                                              For t = 0 To UBound(TabSV)
                                                  If TabSV(t) = "VIDE" Then TabSV(t) = ""
                                                  If (Not LCase(Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne)) _
                                                     Like "*" & LCase(TabSV(t)) & "*") Or nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) = "FFF" Then
                                                     If (valGet = "" And cGet(j, 5) <> cGet(p, 5) And t = UBound(TabSV)) Or (valGet = "" And p = UBound(cGet, 1) And t = UBound(TabSV)) Then
                                                          cnd = False
                                                     End If
                                                 Else
                                                     If cGet(j, 5) = cGet(p, 5) Then findV = True
                                                     valGet = "ok"
                                                 End If
                                             Next t
                                             If cnd = True Then j = j + 1
                                       End If
                                       
                                 ElseIf cGet(j, 3) = "NE CONTIENT PAS" Then
                                         If findV = True And cGet(j, 5) = cGet(j - 1, 5) Then
                                            j = j + 1
                                       Else
                                            findV = False
                                             If InStr(1, cGet(j, 4), ";") <> 0 Then
                                                 TabSV = Split(cGet(j, 4), ";")
                                             Else
                                                 ReDim TabSV(0)
                                                 TabSV(0) = cGet(j, 4)
                                             End If
                                             valGet = ""
                                             p = IIf(j + 1 > UBound(cGet, 1), UBound(cGet), j + 1)
                                              For t = 0 To UBound(TabSV)
                                                  If TabSV(t) = "VIDE" Then TabSV(t) = ""
                                                  If (LCase(Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne)) _
                                                     Like "*" & LCase(TabSV(t)) & "*") Or nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) = "FFF" Then
                                                     If (valGet = "" And cGet(j, 5) <> cGet(p, 5) And t = UBound(TabSV)) Or (valGet = "" And p = UBound(cGet, 1) And t = UBound(TabSV)) Then
                                                          cnd = False
                                                     End If
                                                 Else
                                                     If cGet(j, 5) = cGet(p, 5) Then findV = True
                                                     valGet = "ok"
                                                 End If
                                             Next t
                                             If cnd = True Then j = j + 1
                                       End If
                                 ElseIf cGet(j, 3) = "EGAL A" Then
                                       If findV = True And cGet(j, 5) = cGet(j - 1, 5) Then
                                            j = j + 1
                                       Else
                                            findV = False
                                            If InStr(1, cGet(j, 4), ";") <> 0 Then
                                                TabSV = Split(cGet(j, 4), ";")
                                            Else
                                                ReDim TabSV(0)
                                                TabSV(0) = cGet(j, 4)
                                            End If
                                            valGet = ""
                                             p = IIf(j + 1 > UBound(cGet, 1), UBound(cGet), j + 1)
                                             
                                             For t = 0 To UBound(TabSV)
                                                If TabSV(t) = "VIDE" Then TabSV(t) = ""
                                            '    MsgBox Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, V) & ligne) & " " & TabSV(t)
                                                If (Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne) <> TabSV(t)) Or _
                                                 nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) = "FFF" Then
                                                    If (valGet = "" And cGet(j, 5) <> cGet(p, 5) And t = UBound(TabSV)) Or (valGet = "" And p = UBound(cGet, 1) And t = UBound(TabSV)) Then
                                                       cnd = False
                                                    End If
                                                Else
                                                    If cGet(j, 5) = cGet(p, 5) Then findV = True
                                                    valGet = "ok"
                                                End If
                                            Next t
                                            If cnd = True Then j = j + 1
                                   End If
                                ElseIf cGet(j, 3) = "EGALITE AVEC (SI VALEUR NON VIDE)" Then
                                         If findV = True And cGet(j, 5) = cGet(j - 1, 5) Then
                                            j = j + 1
                                       Else
                                            findV = False
                                           p = IIf(j + 1 > UBound(cGet, 1), UBound(cGet), j + 1)
                                           
                                           If (Len(Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne)) > 0 And _
                                                Len(Wsh.Range(nomColFound(Split(cGet(j, 4), ", ")(0), Split(cGet(j, 4), ", ")(1), SourceFile, v) & ligne)) > 0 And _
                                                Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne) <> _
                                                Wsh.Range(nomColFound(Split(cGet(j, 4), ", ")(0), Split(cGet(j, 4), ", ")(1), SourceFile, v) & ligne)) _
                                                Or nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) = "FFF" Then
                                                  If (cGet(j, 5) <> cGet(p, 5)) Or (p = UBound(cGet, 1)) Then
                                                     cnd = False
                                                 End If
                                         Else
                                             j = j + 1
                                             If cGet(j, 5) = cGet(p, 5) Then findV = True
                                        End If
                                     End If
                                 ElseIf cGet(j, 3) = "EGALITE AVEC" Then
                                      If findV = True And cGet(j, 5) = cGet(j - 1, 5) Then
                                            j = j + 1
                                      Else
                                           p = IIf(j + 1 > UBound(cGet, 1), UBound(cGet), j + 1)
                                           findV = False
                                           If (Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne) <> _
                                                Wsh.Range(nomColFound(Split(cGet(j, 4), ", ")(0), Split(cGet(j, 4), ", ")(1), SourceFile, v) & ligne)) _
                                                Or nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) = "FFF" Then
                                                  If (cGet(j, 5) <> cGet(p, 5)) Or (p = UBound(cGet, 1)) Then
                                                         cnd = False
                                                   End If
                                         Else
                                             j = j + 1
                                             If cGet(j, 5) = cGet(p, 5) Then findV = True
                                        End If
                                     End If
                                ElseIf cGet(j, 3) = "SUPERIEUR A" Then
                                      If findV = True And cGet(j, 5) = cGet(j - 1, 5) Then
                                            j = j + 1
                                      Else
                                        findV = False
                                         If InStr(1, cGet(j, 4), ";") <> 0 Then
                                            TabSV = Split(cGet(j, 4), ";")
                                        Else
                                            ReDim TabSV(0)
                                            TabSV(0) = cGet(j, 4)
                                        End If
                                         valGet = ""
                                         p = IIf(j + 1 > UBound(cGet, 1), UBound(cGet), j + 1)
                                         For t = 0 To UBound(TabSV)
                                           If TabSV(t) = "VIDE" Then TabSV(t) = ""
                                           If (Len(Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne)) > 0 And _
                                            IsNumeric(Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne)) = True And _
                                                Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne) <= _
                                                val(TabSV(t))) _
                                                Or nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) = "FFF" Then
                                                  If (valGet = "" And cGet(j, 5) <> cGet(p, 5) And t = UBound(TabSV)) Or (valGet = "" And p = UBound(cGet, 1) And t = UBound(TabSV)) Then
                                                         cnd = False
                                                 End If
                                             ElseIf IsNumeric(Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne)) = False Then
                                                   If (valGet = "" And cGet(j, 5) <> cGet(p, 5) And t = UBound(TabSV)) Or (valGet = "" And p = UBound(cGet, 1) And t = UBound(TabSV)) Then
                                                         cnd = False
                                                    End If
                                            Else
                                                    If cGet(j, 5) = cGet(p, 5) Then findV = True
                                                    valGet = "ok"
                                            End If
                                        Next t
                                        If cnd = True Then j = j + 1
                                   End If
                                ElseIf cGet(j, 3) = "SUPERIEUR OU EGAL A" Then
                                        If findV = True And cGet(j, 5) = cGet(j - 1, 5) Then
                                            j = j + 1
                                       Else
                                             findV = False
                                             If InStr(1, cGet(j, 4), ";") <> 0 Then
                                                TabSV = Split(cGet(j, 4), ";")
                                            Else
                                                ReDim TabSV(0)
                                                TabSV(0) = cGet(j, 4)
                                            End If
                                             valGet = ""
                                              p = IIf(j + 1 > UBound(cGet, 1), UBound(cGet), j + 1)
                                             For t = 0 To UBound(TabSV)
                                                If TabSV(t) = "VIDE" Then TabSV(t) = ""
                                               If (Len(Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne)) > 0 And _
                                                IsNumeric(Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne)) = True And _
                                                    Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne) < _
                                                    val(TabSV(t))) _
                                                    Or nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) = "FFF" Then
                                                    If (valGet = "" And cGet(j, 5) <> cGet(p, 5) And t = UBound(TabSV)) Or (valGet = "" And p = UBound(cGet, 1) And t = UBound(TabSV)) Then
                                                             cnd = True
                                                     End If
                                                 ElseIf IsNumeric(Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne)) = False Then
                                                         If (valGet = "" And cGet(j, 5) <> cGet(p, 5) And t = UBound(TabSV)) Or (valGet = "" And p = UBound(cGet, 1) And t = UBound(TabSV)) Then
                                                                 cnd = False
                                                         End If
                                                Else
                                                        If cGet(j, 5) = cGet(p, 5) Then findV = True
                                                        valGet = "ok"
                                                         
                                                End If
                                            Next t
                                            If cnd = True Then j = j + 1
                                            
                                       End If
                                ElseIf cGet(j, 3) = "INFERIEUR A" Then
                                     If findV = True And cGet(j, 5) = cGet(j - 1, 5) Then
                                            j = j + 1
                                    Else
                                        findV = False
                                       If InStr(1, cGet(j, 4), ";") <> 0 Then
                                            TabSV = Split(cGet(j, 4), ";")
                                        Else
                                            ReDim TabSV(0)
                                            TabSV(0) = cGet(j, 4)
                                        End If
                                         valGet = ""
                                          p = IIf(j + 1 > UBound(cGet, 1), UBound(cGet), j + 1)
                                         For t = 0 To UBound(TabSV)
                                            If TabSV(t) = "VIDE" Then TabSV(t) = ""
                                            If (Len(Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne)) > 0 And _
                                             IsNumeric(Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne)) = True And _
                                                Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne) >= _
                                                val(TabSV(t))) _
                                                Or nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) = "FFF" Then
                                                If (valGet = "" And cGet(j, 5) <> cGet(p, 5) And t = UBound(TabSV)) Or (valGet = "" And p = UBound(cGet, 1) And t = UBound(TabSV)) Then
                                                         cnd = False
                                                 End If
                                            ElseIf IsNumeric(Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne)) = False Then
                                                 If (valGet = "" And cGet(j, 5) <> cGet(p, 5) And t = UBound(TabSV)) Or (valGet = "" And p = UBound(cGet, 1) And t = UBound(TabSV)) Then
                                                         cnd = False
                                                  End If
                                            Else
                                                    If cGet(j, 5) = cGet(p, 5) Then findV = True
                                                    valGet = "ok"
                                            End If
                                        Next t
                                        If cnd = True Then j = j + 1
                                    End If
                                ElseIf cGet(j, 3) = "INFERIEUR OU EGAL A" Then
                                     If findV = True And cGet(j, 5) = cGet(j - 1, 5) Then
                                            j = j + 1
                                     Else
                                        findV = False
                                       If InStr(1, cGet(j, 4), ";") <> 0 Then
                                            TabSV = Split(cGet(j, 4), ";")
                                        Else
                                            ReDim TabSV(0)
                                            TabSV(0) = cGet(j, 4)
                                        End If
                                         valGet = ""
                                          p = IIf(j + 1 > UBound(cGet, 1), UBound(cGet), j + 1)
                                         For t = 0 To UBound(TabSV)
                                            If TabSV(t) = "VIDE" Then TabSV(t) = ""
                                           
                                            If (Len(Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne)) > 0 And _
                                            IsNumeric(Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne)) = True And _
                                                Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne) > _
                                                val(TabSV(t))) Or nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) = "FFF" Then
                                                If (valGet = "" And cGet(j, 5) <> cGet(p, 5) And t = UBound(TabSV)) Or (valGet = "" And p = UBound(cGet, 1) And t = UBound(TabSV)) Then
                                                        cnd = False
                                                 End If
                                            ElseIf IsNumeric(Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne)) = False Then
                                                If (valGet = "" And cGet(j, 5) <> cGet(p, 5) And t = UBound(TabSV)) Or (valGet = "" And p = UBound(cGet, 1) And t = UBound(TabSV)) Then
                                                         cnd = False
                                                End If
                                            Else
                                                    If cGet(j, 5) = cGet(p, 5) Then findV = True
                                                    valGet = "ok"
                                            End If
                                        Next t
                                        If cnd = True Then j = j + 1
                                    End If
                                ElseIf cGet(j, 3) = "DIFFERENT DE" Then
                                   If findV = True And cGet(j, 5) = cGet(j - 1, 5) Then
                                            j = j + 1
                                   Else
                                       If InStr(1, cGet(j, 4), ";") <> 0 Then
                                            TabSV = Split(cGet(j, 4), ";")
                                        Else
                                            ReDim TabSV(0)
                                            TabSV(0) = cGet(j, 4)
                                        End If
                                         valGet = ""
                                          p = IIf(j + 1 > UBound(cGet, 1), UBound(cGet), j + 1)
                                         For t = 0 To UBound(TabSV)
                                            If TabSV(t) = "VIDE" Then TabSV(t) = ""
                                            If (Wsh.Range(nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) & ligne) = _
                                                (TabSV(t))) Or nomColFound(Split(cGet(j, 2), ", ")(0), Split(cGet(j, 2), ", ")(1), SourceFile, v) = "FFF" Then
                                                If (cGet(j, 5) <> cGet(p, 5)) Or (p = UBound(cGet, 1)) Then
                                                         cnd = False
                                                 End If
                                            End If
                                        Next t
                                        If cnd = True Then j = j + 1
                                      End If
                                End If
                                p = IIf(j > UBound(cGet, 1), UBound(cGet, 1), j)
                                
                                If (IsNumeric(cGet(p, 1)) = True And Len(cGet(p, 1)) > 0 And Len(cGet(p, 3)) = 0 And cnd = True) Or (p = UBound(cGet) And cnd = True) Then
                                  
                                    conditionSDV = Ongls
                                    Exit Function
                                End If
                        Wend
                       
                        cnd = True
                       
                   End If
              Next j

     End With
End Function

Function nomColFound(v1 As Variant, v2 As Variant, SourceFile As Variant, v As Variant)
     If Not colFound.Exists(UCase(v1 & "-" & v2)) Then
            colFound.Add key:=UCase(v1 & "-" & v2), Item:=nomcol(CStr(v1), CStr(v2), SourceFile, v)
            nomColFound = colFound(UCase(v1 & "-" & v2))
    Else
            nomColFound = colFound(UCase(v1 & "-" & v2))
    End If
End Function


Function importAborted(SourceFile As Variant)
    Application.DisplayAlerts = False
    Workbooks(SourceFile).Close False
    ThisWorkbook.Worksheets("DATA").Delete
    Application.DisplayAlerts = True
    Call Moniteur("Importation aborted. Issue with file version (" & SourceFile & ").")
    ThisWorkbook.sheets("HOME").Range("Moniteur").Interior.color = RGB(255, 255, 255)
    ProgressExit
 
End Function














