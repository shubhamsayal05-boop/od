Attribute VB_Name = "Module1"
' Déclaration des fonctions API Windows nécessaires
Declare PtrSafe Function OpenClipboard Lib "user32.dll" (ByVal hwnd As LongPtr) As Long
Declare PtrSafe Function EmptyClipboard Lib "user32.dll" () As Long
Declare PtrSafe Function CloseClipboard Lib "user32.dll" () As Long
 
Sub ViderPressePapiers()

'Ouvre le presse-papiers
If OpenClipboard(0&) Then
    'Vide le presse-papiers
    EmptyClipboard
    'Ferme le presse-papiers
    CloseClipboard
End If
End Sub

