Option Explicit

Public sheet As Worksheet

Private Sub Worksheet_Activate()
  FormReset True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
' Is triggered everytime a cell content changes in the Worksheet

  ' Control if barcode cell was changed and act accordingly
  If Not Intersect(Target, Range("Alterung_Chargenummer")) Is Nothing Then
    If Not Range("Alterung_IsBarcodeEmpty").Value Then
      If Range("Alterung_IsBarcodeValid").Value Then
        If Range("Alterung_Needed").Value Then
          Dim chargeNummer
          chargeNummer = Range("Alterung_Chargenummer").Value

          Select Case Range("Alterung_Phase").Value
          Case chargePhase.Processing ' Process and quality data are missing
            MsgBox "Diese Charge befindet sich noch in Bearbeitung, bitte zuerst die Prozess- und Qualitätsdaten hinzufügen!"
            Sheet1.Activate
            Sheet1.Range("Prozess_Chargenummer").Value = chargeNummer
            Exit Sub
          Case chargePhase.Nacharbeit ' Nacharbeit done and löttest needed
            ' TO DO
            ' Need to split workflow between edit and create?
          Case chargePhase.Alterung
            FormReset False
            'CreateMode_Fill
          Case chargePhase.Done To chargePhase.Scrapped ' Charge is done or scrapped
            FormReset False
            'EditMode_Fill
          Case Else
            MsgBox "Unspecified error with this charge, Phase not handled!"
            FormReset True
          End Select
        Else
          MsgBox "Dieses Teilnummer benötigt keine Löttest!"
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

  Set sheet = Sheet5

  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword

  ' Reset form
  If FullReset Then
    Range("Alterung_Chargenummer").ClearContents
  End If
  Range("Alterung_iO").ClearContents
  Range("Alterung_niO").ClearContents
  Range("Alterung_Mitarbeiter").ClearContents
  Range("Alterung_Kommentar").ClearContents

  Range("Alterung_Chargenummer").Select
  
  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

