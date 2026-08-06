Attribute VB_Name = "Module2"
#If VBA7 Then
    Public Declare PtrSafe Sub sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
#Else
    Public Declare Sub sleep Lib "kernel32" (ByVal Milliseconds As Long)
#End If
' Declare the necessary Windows API functions
Private Declare PtrSafe Function CountClipboardFormats Lib "user32" () As Long

Function fixTheDecimalSeparator(v) As String
    Dim str_v As String
    
    str_v = CStr(v)
    str_v = replace(str_v, ",", ".")
    fixTheDecimalSeparator = str_v
End Function

Sub resetchartsparameter()
    
    With ThisWorkbook.Worksheets("VIERGE")
        .ChartObjects("Graphique_0").Chart.SeriesCollection(1).Border.LineStyle = xlNone
        .ChartObjects("Graphique_1").Chart.SeriesCollection(1).Border.LineStyle = xlNone
        .ChartObjects("Graphique_00").Chart.SeriesCollection(1).Border.LineStyle = xlNone
        .ChartObjects("Graphique_11").Chart.SeriesCollection(1).Border.LineStyle = xlNone
        
        For n = 1 To 13
            .ChartObjects("Graphique P" & n).Chart.SeriesCollection(1).HasDataLabels = True
            .ChartObjects("Graphique P" & n).Chart.SeriesCollection(2).HasDataLabels = True
            For p = 1 To 3
                If p = 1 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(1).Points(p).Format.Fill.ForeColor.RGB = 65535
                If p = 1 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(1).Points(p).Format.Fill.Transparency = 0
                If p = 1 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(1).Points(p).DataLabel.ShowValue = True
                If p = 2 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(1).Points(p).Format.Fill.ForeColor.RGB = 32512
                If p = 2 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(1).Points(p).Format.Fill.Transparency = 0
                If p = 2 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(1).Points(p).DataLabel.ShowValue = True
                If p = 3 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(1).Points(p).Format.Fill.ForeColor.RGB = 255
                If p = 3 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(1).Points(p).Format.Fill.Transparency = 0
                If p = 3 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(1).Points(p).DataLabel.ShowValue = True
                If p = 1 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(2).Points(p).Format.Fill.ForeColor.RGB = 26367
                If p = 1 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(2).Points(p).Format.Fill.Transparency = 0.49
                If p = 1 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(2).Points(p).DataLabel.ShowValue = True
                If p = 2 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(2).Points(p).Format.Fill.ForeColor.RGB = 16777215
                If p = 2 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(2).Points(p).Format.Fill.Transparency = 1
                If p = 2 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(2).Points(p).DataLabel.ShowValue = False
                If p = 3 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(2).Points(p).Format.Fill.ForeColor.RGB = 255
                If p = 3 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(2).Points(p).Format.Fill.Transparency = 0.49
                If p = 3 Then .ChartObjects("Graphique P" & n).Chart.SeriesCollection(2).Points(p).DataLabel.ShowValue = True
            Next
            If n = 3 Then n = 10
        Next
        
        .ChartObjects("RECAPDRI").Chart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = 5287936
        .ChartObjects("RECAPDRI").Chart.SeriesCollection(2).Points(1).Format.Fill.ForeColor.RGB = 65535
        .ChartObjects("RECAPDRI").Chart.SeriesCollection(3).Points(1).Format.Fill.ForeColor.RGB = 255
        .ChartObjects("RECAPDYN").Chart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = 5287936
        .ChartObjects("RECAPDYN").Chart.SeriesCollection(2).Points(1).Format.Fill.ForeColor.RGB = 65535
        .ChartObjects("RECAPDYN").Chart.SeriesCollection(3).Points(1).Format.Fill.ForeColor.RGB = 255
    End With
End Sub
Sub checkDevVersion(cur_version)
    On Error GoTo VersionCheckFailed

    fn = FreeFile
    Open GetDevVersionFilePath() For Input As fn
    Do Until EOF(fn) Or end_zone_s = True
        Line Input #fn, l
        last_version = Split(l, " ")(1)
    Loop
    Close #fn
    last_version_part = Split(last_version, ".")
    cur_version_part = Split(cur_version, ".")
    If CInt(last_version_part(0)) < CInt(cur_version_part(0)) Then
        MsgBox "This version not saved !", vbCritical
        Exit Sub
    ElseIf CInt(last_version_part(0)) > CInt(cur_version_part(0)) Then
        MsgBox "You are not using the last version !", vbCritical
        Exit Sub
    End If
    If CInt(last_version_part(1)) < CInt(cur_version_part(1)) Then
        MsgBox "This version not saved !", vbCritical
        Exit Sub
    ElseIf CInt(last_version_part(1)) > CInt(cur_version_part(1)) Then
        MsgBox "You are not using the last version !", vbCritical
        Exit Sub
    End If
    If CInt(last_version_part(2)) < CInt(cur_version_part(2)) Then
        MsgBox "This version not saved !", vbCritical
        Exit Sub
    ElseIf CInt(last_version_part(2)) > CInt(cur_version_part(2)) Then
        MsgBox "You are not using the last version !", vbCritical
        Exit Sub
    End If
    Exit Sub

VersionCheckFailed:
    On Error Resume Next
    Close #fn
    On Error GoTo 0
End Sub



