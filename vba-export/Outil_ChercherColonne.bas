Attribute VB_Name = "Outil_ChercherColonne"
' Cellule Office

Option Explicit

Function nomcol(ByVal l1 As String, ByVal L2 As String, Optional fichier As Variant, Optional v As Variant) As String

    Dim c As Long
    
    nomcol = "XFD"
    If IsMissing(fichier) = True Then
        With ThisWorkbook.sheets("DATA")
             If Not .Rows(1).Find(What:=l1 & ", " & L2, LookIn:= _
                    xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False) Is Nothing Then
                    nomcol = Split(.Columns(.Rows(1).Find(What:=l1 & ", " & L2, LookIn:= _
                    xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Column).Address(ColumnAbsolute:=False), ":")(1)
              End If
              
'            Workbooks(B).Worksheets(a).Activate
        End With
        
    ElseIf IsMissing(fichier) = False Then
            With Workbooks(fichier).Worksheets("TRIE")
                c = 1
                For c = LBound(v, 2) To UBound(v, 2)
                   If RemoveSpace(CStr(v(1, c))) = RemoveSpace(l1) And RemoveSpace(CStr(v(2, c))) = RemoveSpace(L2) Then
                        nomcol = replace(Left(.Cells(1, c).Address, InStrRev(.Cells(1, c).Address, "$") - 1), "$", "")
                         Exit Function
                   End If
                Next c
                   
            End With
  End If
 
End Function

Function RemoveSpace(st As String)
     Dim i As Long
     RemoveSpace = st
     For i = 0 To Len(RemoveSpace)
         RemoveSpace = replace(RemoveSpace, " ", "")
     Next i
     RemoveSpace = UCase(RemoveSpace)
End Function










