Attribute VB_Name = "modRatingTabs"
Option Explicit

' ================== MACRO PRINCIPALE ==================
Public Sub RT_TrierEtColorerOnglets()

    ' ---- Paramètres ----
    Const RESULT_SHEET_NAME As String = "RATING"
    Const COL_GROUP As Long = 2             ' B
    Const COL_SITU As Long = 4              ' D
    Const ROW_GROUP_START As Long = 23
    Const ROW_SITU_START As Long = 24
    Const KEEP_RESULT_FIRST As Boolean = True
    Const MAX_BLANK_ROWS As Long = 1000     ' large pour ne pas s'arrêter trop tôt

    Dim wb As Workbook, ws As Worksheet, logWs As Worksheet
    Dim lastRow As Long, r As Long
    Dim currentGroup As String
    Dim dictSheetToGroup As Object, dictGroupToColor As Object, dictSeen As Object
    Dim orders As New Collection
    Dim sheetName As String, raw As String
    Dim blanks As Long
    Dim nextColorIdx As Long
    Dim palette As Variant

    On Error GoTo SafeExit
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets(RESULT_SHEET_NAME)

    ' Si la structure du classeur est protégée, Move échouera
    If wb.ProtectStructure Then
        MsgBox "Le classeur est protégé (Structure). Impossible de déplacer les feuilles." & vbCrLf & _
               "Désactive la protection : Révision > Protéger le classeur (décochez 'Structure').", vbCritical
        GoTo SafeExit
    End If

    ' Palette couleurs stable (RGB)
    palette = Array( _
        RGB(79, 129, 189), RGB(192, 80, 77), RGB(155, 187, 89), _
        RGB(128, 100, 162), RGB(75, 172, 198), RGB(247, 150, 70), _
        RGB(146, 208, 80), RGB(0, 176, 240), RGB(112, 48, 160), _
        RGB(255, 192, 0), RGB(0, 112, 192), RGB(192, 0, 0))
    nextColorIdx = 0

    ' Dictionnaires
    Set dictSheetToGroup = CreateObject("Scripting.Dictionary")
    Set dictGroupToColor = CreateObject("Scripting.Dictionary")
    Set dictSeen = CreateObject("Scripting.Dictionary")

    ' Feuille de log (réinitialisation)
    RT_DeleteSheetIfExists "Macro_Log"
    Set logWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    logWs.Name = "Macro_Log"
    logWs.Range("A1:D1").Value = Array("Ligne", "Type", "Info", "Détail")
    logWs.Rows(1).Font.Bold = True

    ' Dernière ligne à lire (selon B ou D)
    lastRow = WorksheetFunction.Max( _
                ws.Cells(ws.Rows.Count, COL_GROUP).End(xlUp).row, _
                ws.Cells(ws.Rows.Count, COL_SITU).End(xlUp).row)
    If lastRow < ROW_GROUP_START Then
        RT_LogLine logWs, 0, "ERROR", "Tableau introuvable", "Vérifie colonnes B/D et lignes 23/24"
        GoTo SafeExit
    End If

    currentGroup = vbNullString
    blanks = 0

    ' ===== Lecture du tableau en ignorant les lignes masquées =====
    For r = ROW_GROUP_START To lastRow

        ' Sauter les lignes masquées (manuelles, outline, filtre)
        If Not RT_RowIsVisible(ws, r) Then GoTo NextRow

        ' --- 1) GROUPE en colonne B (ligne visible) ---
        Dim grpText As String
        grpText = Trim$(CStr(ws.Cells(r, COL_GROUP).Value))
        If grpText <> "" And r >= ROW_GROUP_START Then
            currentGroup = RT_CleanText(grpText)
            If Not dictGroupToColor.Exists(currentGroup) Then
                dictGroupToColor.Add currentGroup, palette(nextColorIdx Mod (UBound(palette) + 1))
                nextColorIdx = nextColorIdx + 1
            End If
            blanks = 0
            ' on continue : une situation peut se trouver sur la même ligne visible (rare)
        End If

        ' --- 2) SITUATION en colonne D (ligne visible) ---
        raw = Trim$(CStr(ws.Cells(r, COL_SITU).Value))
        If raw <> "" And r >= ROW_SITU_START Then

            If currentGroup = "" Then
                currentGroup = "(Sans groupe)"
                If Not dictGroupToColor.Exists(currentGroup) Then
                    dictGroupToColor.Add currentGroup, palette(nextColorIdx Mod (UBound(palette) + 1))
                    nextColorIdx = nextColorIdx + 1
                End If
            End If

            ' Priorité au nom réel via hyperlien
            sheetName = RT_ExtractSheetName(ws.Cells(r, COL_SITU))
            If sheetName = "" Then sheetName = RT_CleanText(raw)

            ' Normalisations usuelles si la feuille n'est pas trouvée du premier coup
            If Not RT_SheetExists(wb, sheetName) Then
                Dim alt As String
                alt = replace(sheetName, "  ", " ")
                If RT_SheetExists(wb, alt) Then
                    sheetName = alt
                Else
                    alt = replace(sheetName, " - ", " ")
                    If RT_SheetExists(wb, alt) Then
                        sheetName = alt
                    Else
                        alt = replace(sheetName, "-", " ")
                        If RT_SheetExists(wb, alt) Then
                            sheetName = alt
                        Else
                            alt = replace(sheetName, " ", "_")
                            If RT_SheetExists(wb, alt) Then
                                sheetName = alt
                            End If
                        End If
                    End If
                End If
            End If

            ' Exclusions explicites
            If UCase$(sheetName) <> "RATING" And UCase$(sheetName) <> "MACRO_LOG" Then
                If RT_SheetExists(wb, sheetName) Then
                    dictSheetToGroup(sheetName) = currentGroup
                    If Not dictSeen.Exists(sheetName) Then
                        orders.Add sheetName
                        dictSeen.Add sheetName, True
                    Else
                        RT_LogLine logWs, r, "INFO", "Doublon situation (ligne visible)", sheetName
                    End If
                Else
                    RT_LogLine logWs, r, "WARN", "Feuille introuvable (ligne visible)", "Nom lu : " & sheetName
                End If
            End If

            blanks = 0
        Else
            ' Incrémenter le compteur de vides uniquement si la ligne est visible et vraiment vide
            If grpText = "" Then
                blanks = blanks + 1
                If blanks >= MAX_BLANK_ROWS Then Exit For
            End If
        End If

NextRow:
    Next r

    If orders.Count = 0 Then
        RT_LogLine logWs, 0, "ERROR", "Aucune situation valide", "Aucun onglet correspondant trouvé"
        GoTo SafeExit
    End If

    ' ===== Coloration des onglets =====
    Dim i As Long, sName As String
    For i = 1 To orders.Count
        sName = orders(i)
        On Error Resume Next
        wb.Worksheets(sName).Tab.color = dictGroupToColor(dictSheetToGroup(sName))
        On Error GoTo 0
    Next i

    ' ===== Réordonnancement sécurisé =====
    RT_ReorderSheetsSafe wb, orders, KEEP_RESULT_FIRST, RESULT_SHEET_NAME

    'MsgBox "Tri et coloration terminés." & vbCrLf & "Voir 'Macro_Log' pour les détails.", vbInformation

SafeExit:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ================== UTILITAIRES (tous préfixés RT_) ==================

Private Function RT_SheetExists(wb As Workbook, sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sName)
    RT_SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function RT_CleanText(s As String) As String
    ' Nettoyage robuste des libellés
    If Len(s) = 0 Then RT_CleanText = "": Exit Function

    ' Supprimer caractères non imprimables standard
    On Error Resume Next
    s = Application.WorksheetFunction.Clean(s)
    On Error GoTo 0

    ' Remplacer espaces insécables & co
    s = replace(s, Chr(160), " ")          ' NBSP
    s = replace(s, ChrW(8239), " ")        ' espace fine insécable
    s = replace(s, ChrW(8201), " ")        ' espace fine
    s = replace(s, vbTab, " ")

    ' Retirer quelques puces / symboles courants
    s = replace(s, "?", "")
    s = replace(s, "•", "")
    s = replace(s, "?", "")
    s = replace(s, "?", "")

    ' Normaliser les tirets / points de suspension
    s = replace(s, ChrW(8211), "-")        ' EN DASH ? -
    s = replace(s, ChrW(8212), "-")        ' EM DASH ? -
    s = replace(s, "…", "...")

    ' Écraser les doubles espaces (2 passes simples)
    s = replace(s, "  ", " ")
    s = replace(s, "  ", " ")

    RT_CleanText = Trim$(s)
End Function

Private Function RT_ExtractSheetName(cell As Range) As String
    On Error GoTo EH
    If cell.Hyperlinks.Count > 0 Then
        Dim subAddr As String, p As Long
        subAddr = cell.Hyperlinks(1).SubAddress
        p = InStr(1, subAddr, "!")
        If p > 1 Then
            RT_ExtractSheetName = replace(Left$(subAddr, p - 1), "'", "")
            Exit Function
        End If
    End If
EH:
    RT_ExtractSheetName = ""
End Function

Private Sub RT_LogLine(logWs As Worksheet, rowNum As Long, kind As String, INFO As String, detail As String)
    With logWs
        Dim nxt As Long
        nxt = .Cells(.Rows.Count, 1).End(xlUp).row + 1
        .Cells(nxt, 1).Value = rowNum
        .Cells(nxt, 2).Value = kind
        .Cells(nxt, 3).Value = INFO
        .Cells(nxt, 4).Value = detail
    End With
End Sub

Private Sub RT_DeleteSheetIfExists(sName As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(sName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Function RT_RowIsVisible(ws As Worksheet, r As Long) As Boolean
    ' True si la ligne r est visible (non masquée / non filtrée)
    On Error Resume Next
    RT_RowIsVisible = Not ws.Rows(r).Hidden
    On Error GoTo 0
End Function

Private Sub RT_ReorderSheetsSafe(wb As Workbook, orders As Collection, _
                                 keepResultFirst As Boolean, resultSheetName As String)
    Dim i As Long
    Dim anchor As Worksheet, sh As Worksheet

    ' Ancre = "RATING" si demandé, sinon première feuille
    If keepResultFirst And RT_SheetExists(wb, resultSheetName) Then
        Set anchor = wb.Worksheets(resultSheetName)
    Else
        Set anchor = wb.Worksheets(1)
    End If

    ' Construire l'ordre final en insérant chaque feuille juste après l'ancre
    For i = 1 To orders.Count
        If RT_SheetExists(wb, orders(i)) Then
            Set sh = wb.Worksheets(orders(i))
            If Not sh Is anchor Then
                If sh.index <> anchor.index + 1 Then
                    sh.Move After:=anchor
                End If
                Set anchor = sh
            End If
        End If
    Next i
End Sub


