Attribute VB_Name = "Feuil023"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
                
                If ActiveSheet.Range("A2:BG2").EntireColumn.Hidden = False Then
                    If Cells(Rows.Count, 13).End(xlUp).row > 6 Then
                       If Not Application.Intersect(Target, Range("M6:" & Cells(TotEventSheet(ActiveSheet.Name), getLastColumnDrivability(ActiveSheet.Name)).Address)) Is Nothing Then
                           createDeleteEvent
                           Cancel = True
                           Application.CommandBars("Del_Event").ShowPopup
                       End If
                    End If
              Else
                     If Cells(Rows.Count, 72).End(xlUp).row > 6 Then
                         If Not Application.Intersect(Target, Range("BT6:" & Cells(TotEventSheet(ActiveSheet.Name), getLastColumnDinamyc(ActiveSheet.Name)).Address)) Is Nothing Then
                            createDeleteEventDyn
                            Cancel = True
                            Application.CommandBars("Del_Event").ShowPopup
                        End If
                    End If
           End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
   If Not Application.Intersect(Target, Range("K11:K19")) Is Nothing Then
     Call CamFull(ActiveSheet.Name, "driv")
  End If

 If Not Application.Intersect(Target, Range("BR11:BR19")) Is Nothing Then
     Call CamFull(ActiveSheet.Name, "dyn")
 End If
End Sub

