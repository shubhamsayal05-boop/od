Attribute VB_Name = "ProgresstionBAR"
Option Explicit

Function ProgressLoad()
    PleaseWait.Show 0
End Function


Function ProgressExit()
   On Error Resume Next
    Unload PleaseWait
    If ERR.Number <> 0 Then ERR.Clear
End Function

Function ProgressTitle(Titre As String)
    PleaseWait.Titre = Titre
    DoEvents
End Function
