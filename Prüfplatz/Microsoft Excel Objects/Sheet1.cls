Option Explicit

Public sheet As Worksheet

Private Sub Worksheet_Activate()
  FormReset False
End Sub

Private Sub FormReset(ByVal FullReset As Boolean)
' Resets the form, deleting all the custom data and moving the selection
' to the barcode Cell

  Set sheet = Sheet1

  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword

  ' Reset form
  If FullReset Then
    Range("Prozess_Passivierung").ClearContents
    Range("Prozess_CuZahlerIst").ClearContents
    Range("Prozess_AgZahlerIst").ClearContents
    Range("Prozess_AgStromIst").ClearContents
  End If
  Range("Prozess_Barcode").ClearContents
  Range("Prozess_Trommel").ClearContents
  Range("Prozess_Kommentar").ClearContents
  Range("Prozess_Mitarbeiter").ClearContents

  Range("Prozess_Barcode").Select

  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub