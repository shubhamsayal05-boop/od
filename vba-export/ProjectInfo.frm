Attribute VB_Name = "ProjectInfo"
Attribute VB_Base = "0{7AFCD9A8-DD23-4B21-9304-5F5B43C7DEFE}{C88F6149-7946-4FCF-B0CE-AC70B5E90B37}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

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


Private Sub Project_AfterUpdate()
    frmNameCode.Show
    Me.Project.text = ThisWorkbook.sheets("home").Range("Project").Value
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



Private Sub UserForm_Activate()
    
    Dim drop As Variant
    Dim i As Integer
    Dim RqOdb As Object
    Dim idc As String
   
    idc = getDbId(ThisWorkbook.Worksheets("Home").Range("idProjects"))
    Set RqOdb = db.GetOdb(val(idc))
'    UserForm_Initialize
    With ThisWorkbook.sheets("HOME")
        Me.Project.text = .Range("project").Value

        'If Len(.Range("mode").Value) > 0 Then Me.Mode.Text = .Range("mode").Value
        If Len(.Range("fuel").Value) > 0 Then Me.Fuel.text = .Range("fuel").Value
        If Len(.Range("area").Value) > 0 Then Me.Area.text = .Range("Area").Value
        If Len(.Range("software").Value) > 0 Then Me.Software.text = .Range("Software").Value
        If Len(.Range("DriveVersion").Value) > 0 Then Me.version.text = .Range("DriveVersion").Value
        'If Len(.Range("Gears").Value) > 0 Then Me.Gearbox.Text = .Range("Gears").Value & " GEARS"
        If db.GetValue("SELECT code FROM dataId WHERE UNIQUENAME= " & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value & " ", RqOdb) <> "" Then
            Me.version.Visible = False
            Me.Label14.Visible = False
        End If
     Me.NbGear.text = .Range("H23").Value
     Me.Gearbox.text = .Range("Gears").Value
        
        If Len(.Range("Milestone").Value) > 0 Then Me.Milestone.text = .Range("Milestone").Value
        
        If Len(.Range("C23").Value) > 0 Then
             For i = 0 To Me.TARGET_VEHICLE.ListCount - 1
                   If InStr(1, "," & .Range("C23").Value & ",", "," & Me.TARGET_VEHICLE.list(i) & ",") Then
                      Me.TARGET_VEHICLE.Selected(i) = True
                   End If
             Next i
        End If
        
        'If .Range("C22").Value <> "" Then Me.TARGET_VEHICLE.Text = .Range("C22").Value
    End With

'    If InStr(1, ThisWorkbook.sheets("RATING").Range("G2").Value, "PARTIAL", vbTextCompare) > 0 Then
'        Me.OptionButton1.Value = True
'    ElseIf InStr(1, ThisWorkbook.sheets("RATING").Range("G2").Value, "FULL", vbTextCompare) > 0 Then
'        Me.OptionButton2.Value = True
'    End If
    
    
    
    drop = db.GetValue("SELECT droopy FROM projet WHERE ID = " & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value & " ")
    If Not isEmpty(drop) Then
        Me.Droopy.Value = drop
    End If
   
End Sub

Private Sub UserForm_Initialize()
    Project.Value = ThisWorkbook.sheets("HOME").Range("Project").Value
    'Mode.Clear
    Me.TARGET_VEHICLE.ColumnWidths = "400"
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
    Dim RechargeP As Boolean
    Dim id_projet As String
    Dim Proj As String
    Dim nomVeh As String
    Dim i As Integer
    Dim idc As String
    Dim RqOdb As Object
    
   nomVeh = ""
    For i = 0 To Me.TARGET_VEHICLE.ListCount - 1
          If Me.TARGET_VEHICLE.Selected(i) = True Then
              nomVeh = IIf(nomVeh = "", Me.TARGET_VEHICLE.list(i), nomVeh & "," & Me.TARGET_VEHICLE.list(i))
          End If
    Next i
    
    'Droopy.Text = "" Or
    If Project.text = "" Or Gearbox.text = "" Or Fuel.text = "" Or Software.text = "" Or Area.text = "" Or nomVeh = "" Or Me.NbGear.text = "" Then
        MsgBox "Project Creating. You must fill all fields.", vbExclamation
        Exit Sub
    
    End If
    
'    If db.GetValue("SELECT UNIQUENAME FROM projet WHERE UNIQUENAME= '" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value & "' ") <> "" Then
'        MsgBox "Project Creating. Projet Exists.", vbCritical
'        Exit Sub
'    End If
              
     Proj = db.GetValue("SELECT UniqueName FROM projet WHERE UniqueName= " & Chr(34) & Droopy.Value & "_" & ThisWorkbook.Worksheets("HOME").Range("AT32").Value & "_" & Gearbox.Value & "_" & Fuel.Value & "_" & Milestone.Value & "_" & Area.Value & "_" & "PREMIUM" & "_" & Software.Value & "_" & nomVeh & "_" & version.Value & Chr(34) & " And ID<> " & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value)
    If Proj <> "" Then
        MsgBox "Error : Cannot registred project" & vbCrLf & "You project whith this 'Name/code' and this version already exists", vbCritical
        Exit Sub
    End If
    
    'Application.ScreenUpdating = False

    ProjectInfo.hide
   
    With ThisWorkbook.sheets("HOME")
        .Range("Project") = Project.Value
        .Range("Fuel") = Fuel.Value
         .Range("Gears") = Gearbox.Value
        .Range("Area") = Area.Value
        RechargeP = False
        
        .Range("Targets") = "PREMIUM"
        .Range("Milestone") = Milestone.Value
       
        .Range("Software") = Software.Value
        .Range("C23") = nomVeh
        .Range("H23") = NbGear.Value
        .Activate
        
        ThisWorkbook.sheets("HOME").Range("AT32").Locked = False
        
     
        id_projet = db.GetValue("SELECT ID FROM projet WHERE ID= " & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value & " ")
        Dim id_soft As Variant
        id_soft = db.GetValue("SELECT id FROM milestone WHERE software= '" & Software.Value & "' ")
         .Range("DriveVersion") = version.Value
        

  db.Execute "UPDATE projet SET uniquename=" & Chr(34) & Droopy.Value & "_" & ThisWorkbook.Worksheets("HOME").Range("AT32").Value & "_" & Gearbox.Value & "_" & Fuel.Value & "_" & Milestone.Value & "_" & Area.Value & "_" & "PREMIUM" & "_" & id_soft & "_" & nomVeh & "_" & version.Value & Chr(34) & _
  ", version='" & version.Value & "', droopy='" & Droopy.Value & "', target_vehicle='" & nomVeh & "', gears='" & Gearbox.Value & "', energy='" & Fuel.Value & "', milestone='" & Milestone.Value & "', aera='" & Area.Value & "' , software='" & Me.Software.Value & "',  NbGear='" & NbGear.Value & "' WHERE ID= " & id_projet & " "
    
    idc = getDbId(id_projet)
    
    If idc <> "" Then
         Set RqOdb = db.GetOdb(val(idc))
         db.Execute "UPDATE projet SET uniquename=" & Chr(34) & Droopy.Value & "_" & ThisWorkbook.Worksheets("HOME").Range("AT32").Value & "_" & Gearbox.Value & "_" & Fuel.Value & "_" & Milestone.Value & "_" & Area.Value & "_" & "PREMIUM" & "_" & id_soft & "_" & nomVeh & "_" & version.Value & Chr(34) & _
        ", version='" & version.Value & "', droopy='" & Droopy.Value & "', target_vehicle='" & nomVeh & "', gears='" & Gearbox.Value & "', energy='" & Fuel.Value & "', milestone='" & Milestone.Value & "', aera='" & Area.Value & "' , software='" & Me.Software.Value & "',  NbGear='" & NbGear.Value & "' WHERE ID= " & id_projet & " ", RqOdb
    End If
         db.CloseSudbConn
         MsgBox "Project Creating. Project successfully modified.", vbInformation, "ODRIV"
         If RechargeP = True Then Call Recharge
        ThisWorkbook.sheets("HOME").Range("AT32").Locked = True
        
    End With

    ThisWorkbook.sheets("TARGETS").Visible = False
  

End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub


Function Recharge()
    
End Function

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









