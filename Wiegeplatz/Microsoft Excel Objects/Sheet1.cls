Option Explicit

Public sheet As Worksheet
Private barCodeCell, barCodeValueCell, chargeWeightCell, anlageCell, isTeileInDBCell As Range
Private isFormValidCell, isBarcodeValidCell, isBarcodeEmptyCell, kupferSchichtCell As Range
Private silberSchichtCell As Range
Private materialNummerInDB As Boolean

Private Sub Worksheet_Activate()
  FormReset True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
' Is triggered everytime a cell content changes in the Worksheet

  ' Define cell ranges
  Set barCodeCell = Range("NeueCharge_Barcode")
  Set isTeileInDBCell = Range("NeueCharge_IsTeileInDatabase")
  Set isFormValidCell = Range("NeueCharge_IsFormValid")
  Set isBarcodeValidCell = Range("NeueCharge_IsBarcodeValid")
  Set isBarcodeEmptyCell = Range("NeueCharge_IsBarcodeEmpty")
  
  ' Control if barcode cell was changed and act accordingly
  If Not Intersect(Target, barCodeCell) Is Nothing Then
    If Not isBarcodeEmptyCell.Value Then
      If isBarcodeValidCell.Value Then
        If isTeileInDBCell Then
          FormReset False
          Form_Fill
        Else
          If MsgBox("Teilenummer nicht vorhanden!" & vbNewLine & "Jetzt hinzufügen?", vbInformation + vbYesNo + vbDefaultButton1, "ACHTUNG") = vbYes Then
            Dim barcode
            barcode = barCodeCell.Value
            Sheet4.Activate
            Sheet4.Range("NeueTeile_Barcode").Value = barcode
            Exit Sub
          End If
          FormReset True
        End If
      Else
        MsgBox "Ungültiger barcode!"
        FormReset True
      End If
    End If
  End If
  
  ' Enable or disable the save button
  NeueCharge_SpeichernBtn.Enabled = False
  NeueCharge_SpeichernBtn.BackColor = vbButtonFace
  If isFormValidCell.Value Then
    NeueCharge_SpeichernBtn.Enabled = True
    NeueCharge_SpeichernBtn.BackColor = RGB(0, 192, 0)
  End If
End Sub

Private Sub FormReset(ByVal BarcodeReset As Boolean)
' Resets the form, deleting all the custom data and moving the selection
' to the barcode cell

  Set sheet = Sheet1
  
  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword
  
  ' Reset form
  If BarcodeReset Then
    Set barCodeCell = Range("NeueCharge_Barcode")
    barCodeCell.ClearContents
  End If
  
  Set chargeWeightCell = Range("NeueCharge_Gewicht")
  chargeWeightCell.ClearContents
  
  Set anlageCell = Range("NeueCharge_Anlage")
  anlageCell.ClearContents
  
  Set kupferSchichtCell = Range("NeueCharge_CuSchicht")
  kupferSchichtCell.ClearContents
  
  Set silberSchichtCell = Range("NeueCharge_AgSchicht")
  silberSchichtCell.ClearContents
  
  barCodeCell.Select
  
  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub Form_Fill()
' Compile the form with the data scanned from the Kanban Karte and the data
' retrieved from the database

  ' Copy "Kupferschicht"
  Set kupferSchichtCell = Range("NeueCharge_CuSchicht")
  If Range("I13").Value <> "" Then
    If CLng(Range("I13").Value) <> 0 Then
      kupferSchichtCell.Value = Range("I13").Value
    Else
      kupferSchichtCell.Value = "NEIN"
    End If
  End If

  ' Copy "Silberschicht"
  Set silberSchichtCell = Range("NeueCharge_AgSchicht")
  If Range("I14").Value <> "" Then
    If CLng(Range("I14").Value) <> 0 Then
      silberSchichtCell.Value = Range("I14").Value
    End If
  End If
  
  ' Copy "Standardanlage"
  Set anlageCell = Range("NeueCharge_Anlage")
  anlageCell.Value = Range("I16").Value
End Sub

Private Sub NeueCharge_SpeichernBtn_Click()
' Save workflow triggered by Save button.
' This form will always create a new Chargenummer and add it to the
' Production database
'
' TASKS
' [x] Define the charges database structure
' [ ] Add Mitarbeiter- and Kommentarfield
' [ ] Check if the Database file is already open
' [ ] Wait for the Database file to close before saving
' [ ] Write data to the database
' [ ] Update all the linked Query in the file after saving
' [ ] Show saving message
  SaveCharge
End Sub

Private Sub SaveCharge()
' Reads the form data and compile an array to write. Then opens the database file
' in hidden mode and write the data in the database. It will always produce a new
' row with a unique ChargeNummer.

  ' Show saving messaage
  Application.StatusBar = "Datenbank wird geöffnet, bitte warten ..."

  ' Definition of variables

  ' Preparing data for write

  ' Check if database is open

  ' Opening of database file
  Application.StatusBar = "Wird gespeichert ..."

  ' Setting of targets

  ' Write data in database

  ' Save and close the database

  ' Refresh all data connections & queries

  ' Show confirmation message
  Application.StatusBar = "Neue Charge für Material " & Range("E7").Value & " erfolgreich gespeichert"
  Notify Range("E7").Value, "Charge erfolgreich hinzugefügt", "Info", 10

  ' Reset the form
  'FormReset True
  MsgBox "It works!"
End Sub

Private Sub NeueCharge_AbbrechenBtn_Click()
  FormReset True
End Sub
