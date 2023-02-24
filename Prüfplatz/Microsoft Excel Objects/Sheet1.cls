Option Explicit

Public sheet As Worksheet

Private Sub Worksheet_Activate()
  FormReset False
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
' Is triggered everytime a cell content changes in the Worksheet

  ' Control if barcode cell was changed and act accordingly
  If Not Intersect(Target, Range("Prozess_Barcode")) Is Nothing Then
    If Not Range("Prozess_IsBarcodeEmpty").Value Then
      If Range("Prozess_IsBarcodeValid").Value Then
        If Range("Prozess_IsChargeInDB").Value Then
          If Range("Prozess_Phase").Value = 0 Then
            FormReset False
            CreateMode_Fill
          Else
            FormReset False
            EditMode_Fill
          End If
        Else
          MsgBox "Die Charge wurde nicht in der Datenbank gefunden!"
          FormReset True
        End If
      Else
        MsgBox "Ungültiger barcode!"
        FormReset True
      End If
    End If
  End If
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
    Range("Prozess_Barcode").ClearContents
  End If
  Range("Prozess_Trommel").ClearContents
  Range("Prozess_Kommentar").ClearContents
  Range("Prozess_Mitarbeiter").ClearContents

  Range("Prozess_Barcode").Select

  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub CreateMode_Fill()
  MsgBox "Create mode aktiv"
End Sub

Private Sub EditMode_Fill()
  MsgBox "Edit mode aktiv"
End Sub
