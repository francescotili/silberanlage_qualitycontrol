Option Explicit

Public sheet As Worksheet

Private Sub FormReset(ByVal FullReset As Boolean)
' Resets the form, deleting all the custom data and moving the selection
' to the barcode Cell

  Set sheet = Sheet1

  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword

  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub