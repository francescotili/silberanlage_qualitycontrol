Option Explicit

Public sheet As Worksheet

Private Sub Worksheet_Activate()
  FormReset
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
' Is triggered everytime a cell content changes in the Worksheet

  ' Control if barcode cell was changed and act accordingly
  If Not Intersect(Target, Range("Prufplatz_Chargenummer")) Is Nothing Then
    If Not Range("Prufplatz_IsBarcodeEmpty").Value Then
      If Range("Prufplatz_IsBarcodeValid").Value Then
        If Range("Prufplatz_IsChargeInDB").Value Then
          Dim chargeNummer
          chargeNummer = Range("Prufplatz_Chargenummer").Value

          Select Case Range("Prufplatz_Phase")
          Case chargePhase.Processing ' Process and quality data are missing
            Sheet1.Activate
            Sheet1.Range("Prozess_Chargenummer").Value = chargeNummer
            Exit Sub
          Case chargePhase.Nacharbeit ' Nacharbeit done, choose based on nacharbeit type
            'Select Case nacharbeit_type
            'SheetX.Activate
            'SheetX.Range("").Value = xx
          Case chargePhase.Alterung ' Alterung done
            Sheet5.Activate
            Sheet5.Range("Alterung_Chargenummer").Value = chargeNummer
          Case chargePhase.Done To chargePhase.Scrapped ' Charge is done or scrapped
            Sheet6.Activate
            Sheet6.Range("Info_Chargenummer").Value = chargeNummer
          Case Else
            MsgBox "Unspecified error"
            FormReset
          End Select
        Else
          MsgBox "Die Charge wurde nicht in der Datenbank gefunden!"
          FormReset
        End If
      Else
        MsgBox "Ungültiger barcode!"
        FormReset
      End If
    End If
  End If
End Sub

Private Sub FormReset()
' Resets the form, deleting all the custom data and moving selection
' to the barcode cell
  Set sheet = Sheet2
  
  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword

  ' Reset form
  Range("Prufplatz_Chargenummer").ClearContents

  Range("Prufplatz_Chargenummer").Select
  
  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub
