Attribute VB_Name = "New_Project"
Attribute VB_Base = "0{6BAEF7D1-3BF4-4792-B567-01C36FC57B8A}{A58BEA99-99A4-4269-BCA8-1CD2CE61C8F9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

' Cellule Office
Option Explicit



Private Sub OptionButton1_Click()
    If Me.OptionButton1.Value Then
'        ThisWorkbook.sheets("RATING").Range("G2").Value = "PARTIAL Driveability Index"
    Else
'        ThisWorkbook.sheets("RATING").Range("G2").Value = "FULL Driveability Index"
    End If
End Sub

Private Sub OptionButton2_Click()
    If Me.OptionButton1.Value Then
'        ThisWorkbook.sheets("RATING").Range("G2").Value = "PARTIAL Driveability Index"
    Else
'        ThisWorkbook.sheets("RATING").Range("G2").Value = "FULL Driveability Index"
    End If
End Sub

Private Sub Project_Enter()
    frmNameCode.Show
    Me.Project.text = ThisWorkbook.sheets("home").Range("Project").Value
    Me.Software.SetFocus
End Sub

Private Sub Software_Change()
   
        
  Dim c As Range
    With ThisWorkbook.Worksheets("CONFIGURATIONS")
        Set c = .Range("MILESTONE")
        Set c = c.Offset(1, 0)
        While c.Value <> ""
            If c.Value = Software.Value Then
                Me.Milestone.Value = c.Offset(0, 1)
               
            End If
             Set c = c.Offset(1, 0)
        Wend
       
   End With

End Sub


Private Sub UserForm_Initialize()
    Project.Value = ThisWorkbook.sheets("HOME").Range("Project").Value
    Me.TARGET_VEHICLE.ColumnWidths = "400"
    Droopy.Value = ""
    Fuel.Clear
    Gearbox.Clear
    Area.Clear
    Milestone.Value = ""

    RemplilListe ("ENGINE")
    RemplilListe ("GEARBOX")
    RemplilListe ("VERSION")
    RemplilListe ("AREA")
    
    RemplilListe ("VEHICLE")
    RemplilListe ("MILESTONE")
    RemplilListe ("NBGEAR")
    
    Me.TARGET_VEHICLE.Left = 384
    Me.TARGET_VEHICLE.Top = 108
    
End Sub


Private Sub ApplyButton_Click()
    
    Dim Proj As Variant
    Dim i As Integer
    Dim nomVeh As String
    Dim getDB As String
    Dim id As String
    Dim RqOdb As Object
    
   nomVeh = ""
    For i = 0 To Me.TARGET_VEHICLE.ListCount - 1
          If Me.TARGET_VEHICLE.Selected(i) = True Then
              nomVeh = IIf(nomVeh = "", Me.TARGET_VEHICLE.list(i), nomVeh & "," & Me.TARGET_VEHICLE.list(i))
          End If
    Next i
    
    If Me.NbGear.text = "" Or Droopy.text = "" Or Project.text = "" Or Gearbox.text = "" Or Fuel.text = "" Or Software.text = "" Or Area.text = "" Or nomVeh = "" Then
        MsgBox "Project Creating. You must fill all fields.", vbExclamation
        Exit Sub
    
    End If
     
     
    Proj = db.GetValue("SELECT UniqueName FROM projet WHERE UniqueName= " & Chr(34) & Droopy.Value & "_" & Project.Value & "_" & Gearbox.Value & "_" & Fuel.Value & "_" & Milestone.Value & "_" & Area.Value & "_" & "PREMIUM" & "_" & Me.Software.Value & "_" & nomVeh & "_" & version.Value & Chr(34))
    If Proj <> "" Then
        MsgBox "Error : Cannot registred project" & vbCrLf & "You project whith this 'Name/code' and this version already exists", vbCritical
        Exit Sub
    End If
    
    'Application.ScreenUpdating = False

    Call Moniteur("New Project started.")
    
    New_Project.hide

    With ThisWorkbook.sheets("HOME")
        .Range("Project") = Project.Value
        .Range("Fuel") = Fuel.Value

       .Range("Gears") = Gearbox.Value
       .Range("DriveVersion") = version.Value
        .Range("Area") = Area.Value
       

        .Range("Targets") = "PREMIUM"
        .Range("Milestone") = Milestone.Value
        '.Range("DriveVersion") = Version.Value
        .Range("Software") = Software.Value
        .Range("C23") = nomVeh
        .Range("H23") = NbGear.Value
        
        .Activate
      
       
        ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value = Droopy.Value & "_" & Project.Value & "_" & Gearbox.Value & "_" & Fuel.Value & "_" & Milestone.Value & "_" & Area.Value & "_" & "PREMIUM" & "_" & Me.Software & "_" & nomVeh & "_" & version.Value
        'db.Execute "INSERT INTO projet (code, gears, energy, priority, milestone, aera, target, software, type) VALUES ('" & Project.Value & "', '" & Left(Gearbox.Value, 1) & "', '" & Fuel.Value & "', '" & Prestation.Value & "' ,'" & Milestone.Value & "', '" & Area.Value & "', 'PREMIUM', '" & id_soft & "', '" & typeProject & "' )"
        
        getDB = dbNameBalancer
        db.Execute "INSERT INTO projet (droopy, code, gears, energy, milestone, aera, target, software, target_vehicle, version, uniqueName, NbGear) VALUES ('" & Droopy.Value & "', '" & Project.Value & "', '" & Gearbox.Value & "', '" & Fuel.Value & "', '" & Milestone.Value & "', '" & Area.Value & "', 'PREMIUM', '" & Me.Software.Value & "', '" & nomVeh & "', '" & version.Value & "', '" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value & "', '" & NbGear.Value & "')"
        
         
         id = db.GetValue("Select Max(Id) from projet")
         db.Execute "INSERT INTO projet" & db.AnneeEnCours & " (db_name, code) VALUES ('" & getDB & "', " & id & ")"
        
                      
        Set RqOdb = db.GetOdb(val(Right(getDB, 1)))
        db.Execute "INSERT INTO projet (droopy, code, gears, energy, milestone, aera, target, software, target_vehicle, version, uniqueName, NbGear, id) VALUES ('" & Droopy.Value & "', '" & Project.Value & "', '" & Gearbox.Value & "', '" & Fuel.Value & "', '" & Milestone.Value & "', '" & Area.Value & "', 'PREMIUM', '" & Me.Software.Value & "', '" & nomVeh & "', '" & version.Value & "', '" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value & "', '" & NbGear.Value & "', " & id & ")", RqOdb
       
        ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value = db.GetValue("Select Id from projet where uniquename=" & Chr(34) & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value & Chr(34))
        ThisWorkbook.Worksheets("HOME").Range("idProjects") = ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value
    End With
    
    db.CloseSudbConn
    MsgBox "Project Creating. Project successfully registered.", vbInformation, "ODRIV"

    ThisWorkbook.sheets("TARGETS").Visible = False

   
End Sub


Private Sub CancelButton_Click()
    Unload Me
End Sub

Function RemplilListe(Rg As String)
    Dim c As Range
     With ThisWorkbook.Worksheets("CONFIGURATIONS")
         Set c = .Range(Rg)
         Set c = c.Offset(1, 0)
         While c.Value <> ""
             If Rg = "VERSION" Then
                 version.AddItem "V" & c.Value
             ElseIf Rg = "VEHICLE" Then
                 TARGET_VEHICLE.AddItem c.Value
             ElseIf Rg = "MILESTONE" Then
                 Software.AddItem c.Value
             ElseIf Rg = "AREA" Then
                 Area.AddItem c.Value
             ElseIf Rg = "ENGINE" Then
                 Fuel.AddItem c.Value
             ElseIf Rg = "GEARBOX" Then
                 Gearbox.AddItem c.Value
             ElseIf Rg = "NBGEAR" Then
                 NbGear.AddItem c.Value
             End If
             Set c = c.Offset(1, 0)
         Wend
    End With
End Function



