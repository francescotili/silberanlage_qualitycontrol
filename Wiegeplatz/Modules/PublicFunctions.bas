Private Function PrufzeichenRechnen(text As String) As String
    Dim i, j, checksumme As Integer
    Dim code128 As String
    Dim fehlzeichen As Boolean
    Dim zeichensatz As Variant
    zeichensatz = Array("ß", "!", Chr(34), "#", "$", "%", "&", "'", "(", ")", "*", "+", ",", "-", ".", "/", _
        "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
        ":", ";", "<", "=", ">", "?", "@", _
        "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", _
        "[", "\", "]", "^", "_", "`", _
        "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", _
        "{", "|", "}", "~", _
        "Ã", "Ä", "Å", "Æ", "Ç", "È", "É", "Ê", "Ë", "Ì", "Í", "Î")
    code128 = ""
    checksumme = 104
    
    ' Prufzeichen berechnen
    For i = 1 To Len(text)
    fehlzeichen = True
        For j = 0 To 94
            If (Mid(text, i, 1) = zeichensatz(j)) Then
                fehlzeichen = False
                checksumme = checksumme + (i * j)
                Exit For
            End If
        Next j
        If fehlzeichen = True Then
            MsgBox "Das Zeichen " & Mid(text, i, 1) & " kann nicht dargestellt werden.", vbCritical, "Barcode Generation function"
            Exit Function
        End If
    Next i
    
    ' Rest ermitteln
    checksumme = checksumme Mod 103

    ' Return Ergebniss
    PrufzeichenRechnen = zeichensatz(checksumme)
End Function

Public Function BarCode128(barcodetext As String) As String
    BarCode128 = "Ì" & barcodetext & PrufzeichenRechnen(CStr(barcodetext)) & "Î"
End Function