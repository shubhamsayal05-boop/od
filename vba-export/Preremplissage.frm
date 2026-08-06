Attribute VB_Name = "Preremplissage"
Attribute VB_Base = "0{5DD06A31-5617-43FE-BC5C-9D4E6167EE1D}{B0A97467-9285-47A2-AAC9-EDB3733C1362}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



Option Explicit
Private fieldsR(15) As String
Private Sub ComboBox3_Change()
    getUser
End Sub

Private Sub ComboBox3_Enter()
    initList
End Sub

Private Sub CommandButton2_Click()
   If isEmptyFields = True Then
        MsgBox "Remplir Tous Les Champs", vbCritical, "Odriv"
    Else
         Call Report.REPORTC(fieldsR)
    End If
End Sub


Private Sub CommandButton3_Click()
 If isEmptyFields = True Then
        MsgBox "Remplir Tous Les Champs", vbCritical, "Odriv"
Else
     Call ViderPressePapiers
     Call Report_PPT.REPORTC_PPT(fieldsR)
End If
End Sub


Private Sub Stage_Change()

End Sub

Private Sub UserForm_Initialize()
    Me.Project = ThisWorkbook.Worksheets("HOME").Range("Project").Value
    Me.From.Locked = True
    Me.Email.Locked = True
    Me.Tel.Locked = True
End Sub

Sub initList()
        Dim r As Long
        Dim i As Integer
        Application.EnableEvents = False
        Me.ComboBox3.RowSource = ""
        With ThisWorkbook.sheets("Utilisateurs")
                r = .Range("A65000").End(xlUp).row
                 For i = 2 To r
                            If Len(.Cells(i, 1)) > 0 Then
                               Me.ComboBox3.AddItem .Cells(i, 1)
                            End If
                 Next i
        End With
        
        
        Application.EnableEvents = True

End Sub

Sub getUser()
        Dim r As Long
        Dim i As Integer
        Application.EnableEvents = False
        Me.From.Value = ""
        Me.Tel.Value = ""
        Me.Email.Value = ""
        With ThisWorkbook.sheets("Utilisateurs")
                r = .Range("A65000").End(xlUp).row
                 For i = 2 To r
                            If .Cells(i, 1) = Me.ComboBox3.Value Then
                                 Me.From.Value = .Cells(i, 2)
                                 Me.Tel.Value = .Cells(i, 3)
                                 Me.Email.Value = .Cells(i, 4)
                                 Me.DepTV.Value = .Cells(i, 5)
                                 Exit Sub
                            End If
                 Next i
        End With
        
        
        Application.EnableEvents = True

End Sub


Function isEmptyFields() As Boolean
    Dim Fields(15) As Object
    Dim fieldsName(15) As String
    Dim i As Integer
    
    isEmptyFields = False
    Set Fields(1) = Me.Project
    Set Fields(2) = Me.Domain
    Set Fields(3) = Me.Standard
    Set Fields(4) = Me.Stage
    Set Fields(5) = Me.Goal
    Set Fields(6) = Me.Site
    Set Fields(7) = Me.CarNumber
    Set Fields(8) = Me.LocDateTest
    Set Fields(9) = Me.Climate
    Set Fields(10) = Me.AcStatus
    Set Fields(11) = Me.VehicleOption
    Set Fields(12) = Me.From
    Set Fields(13) = Me.Tel
    Set Fields(14) = Me.Email
    Set Fields(15) = Me.DepTV
    
     fieldsName(1) = "ProjectHead;Signet1;tetTabl1_1"
     fieldsName(2) = "Domains"
     fieldsName(3) = "Standars"
     fieldsName(4) = "Stages"
     fieldsName(5) = "Goals"
     fieldsName(6) = "Signet7"
     fieldsName(7) = "tetTabl1_2"
     fieldsName(8) = "LocDatTest"
     fieldsName(9) = "ClimCond"
     fieldsName(10) = "AcStatus"
     fieldsName(11) = "vehOption"
     fieldsName(12) = "SendFrom"
     fieldsName(13) = "TelFrom"
     fieldsName(14) = "MailFrom"
     fieldsName(15) = "DepTV"
        
    For i = 1 To UBound(Fields)
        If Len(Fields(i).Value) = 0 Then
             isEmptyFields = True
             Exit Function
        Else
            fieldsR(i) = Fields(i).Value & "#" & fieldsName(i)
        End If
    Next i
    
End Function



