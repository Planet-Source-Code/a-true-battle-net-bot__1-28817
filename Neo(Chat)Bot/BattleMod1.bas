Attribute VB_Name = "BattleMod1"

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Username As String
Public Password As String
Public ConnectedToServer As Boolean
Public RealmServer As String
Public Sub removelitem(List2 As ListBox)
    On Error GoTo ErrHand
    LastLI = (List2.ListIndex)
    List2.RemoveItem (List2.ListIndex)
    List2.ListIndex = LastLI
ErrHand:


    If Err.Number = 0 Then
    ElseIf Err.Number = 380 Then
        List2.ListIndex = LastLI - 1
    End If
End Sub
Public Function Encrypt(Text As String) As String
    Dim x$, X2$, i%, Boo%, Boo2$, PATorJK$
    x$ = " ?!@#$%^&*()_+|0123456789abcdefghijklmnopqrstuvwxyz.,-~ABCDEFGHIJKLMNOPQRSTUVWXYZ¿¡²³ÀÁÂÃÄÅÒÓÔÕÖÙÛÜàáâãäåØ¶§Ú¥"
    X2$ = " ¿¡@#$%^&*()_+|9876543210.bÁd-Ã*h^jk$m#ÒÓ]ÔtÖ+vw|Üz.,-~~'áâãFRgäJ¥å\/Ø¶zQR§uÚV0X¥Z?!23acefinoprstuxyBCDEILOPSUY"
    For i% = 1 To Len(Text)
        Boo% = InStr(x$, Mid(Text, i%, 1))
        If Not Boo% = 0 Then
            Boo2$ = Mid(X2$, Boo%, 1)
            PATorJK$ = PATorJK$ + Boo2$
        End If
    Next
    Encrypt = PATorJK$
End Function

