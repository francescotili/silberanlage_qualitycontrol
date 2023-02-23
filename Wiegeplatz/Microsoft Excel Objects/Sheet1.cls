Option Explicit

Public sheet As Worksheet
Private barCodeCell, barCodeValueCell, isTeileInDBCell As Range
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
  
  Range("NeueCharge_Gewicht").ClearContents
  Range("NeueCharge_Anlage").ClearContents
  Range("NeueCharge_CuSchicht").ClearContents
  Range("NeueCharge_AgSchicht").ClearContents
  Range("NeueCharge_Kommentar").ClearContents
  
  barCodeCell.Select
  
  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub Form_Fill()
' Compile the form with the data scanned from the Kanban Karte and the data
' retrieved from the database
  
  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword

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
  Range("NeueCharge_Anlage").Value = Range("I16").Value
  
  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub NeueCharge_SpeichernBtn_Click()
' Save workflow triggered by Save button.
' This form will always create a new Chargenummer and add it to the
' Production database
'
' TASKS
' [x] Define the charges database structure
' [x] Add Mitarbeiter- and Kommentarfield
' [x] Check if the Database file is already open
' [x] Wait for the Database file to close before saving
' [x] Use all named ranges for data preparation
' [x] Write data to the database
' [x] Update all the linked Query in the file after saving
' [x] Show saving message
  SaveCharge
End Sub

Private Sub SaveCharge()
' Reads the form data and compile an array to write. Then opens the database file
' in hidden mode and write the data in the database. It will always produce a new
' row with a unique ChargeNummer.

  ' Show saving messaage
  Application.StatusBar = "Datenbank wird geöffnet, bitte warten ..."

  ' Definition of variables
  Dim app As New excel.Application
  Dim wbook As excel.Workbook
  Dim sheet As Worksheet
  Dim table As ListObject
  Dim chargeData As Variant
  Dim newRow As ListRow
  Dim chargeNummer As String
  
  ' Freeze the Chargenummer
  chargeNummer = Range("NeueCharge_Chargenummer").Value

  ' Check if database is open
  Do While IsDatabaseOpen = True
    Application.StatusBar = "Datenbank beschäftigt, bitte warten ..."
    Application.Wait (Now + TimeValue("0:00:02"))
  Loop

  ' Opening of database file
  app.Visible = False ' Visibile is False by default, so this should not be necessary
  app.DisplayAlerts = False
  Set wbook = app.Workbooks.Add(GDatabasePath)
  Application.StatusBar = "Wird gespeichert ..."

  ' Setting of targets
  Set sheet = wbook.Worksheets(1) ' Main Database
  Set table = sheet.ListObjects.Item(GQSilberDB_TableName)

  ' Write data in database
  table.ListRows.Add Position:=1
  With table
    .DataBodyRange(1, QSilberDB_Col.Chargenummer).Value = chargeNummer
    .DataBodyRange(1, QSilberDB_Col.Datum).Value = Range("NeueCharge_Datum").Value
    .DataBodyRange(1, QSilberDB_Col.Schicht).Value = Range("NeueCharge_Schicht").Value
    .DataBodyRange(1, QSilberDB_Col.Annahme_Mitarbeiter).Value = Range("NeueCharge_Mitarbeiter").Value
    .DataBodyRange(1, QSilberDB_Col.Annhame_Kommentar).Value = Range("NeueCharge_Kommentar").Value
    ' I7 is Materialnummer for DB
    .DataBodyRange(1, QSilberDB_Col.Materialnummer).Value = Range("I7").Value
    .DataBodyRange(1, QSilberDB_Col.Gewicht_netto).Value = Range("NeueCharge_Nettogewicht").Value
    .DataBodyRange(1, QSilberDB_Col.CuSchicht_soll).Value = Range("NeueCharge_CuSchicht").Value
    .DataBodyRange(1, QSilberDB_Col.AgSchicht_soll).Value = Range("NeueCharge_AgSchicht").Value
    .DataBodyRange(1, QSilberDB_Col.Auftragsnummer).Value = Range("NeueCharge_Auftrag").Value
    .DataBodyRange(1, QSilberDB_Col.Füllmenge).Value = Range("NeueCharge_FullmengeAFO").Value
    .DataBodyRange(1, QSilberDB_Col.Stückzahl).Value = Range("NeueCharge_Stuckzahl").Value
    .DataBodyRange(1, QSilberDB_Col.AgBedarf_soll).Value = Range("NeueCharge_SilberBedarfsoll").Value
    .DataBodyRange(1, QSilberDB_Col.Anlage).Value = Range("NeueCharge_Anlage").Value
    .DataBodyRange(1, QSilberDB_Col.CuWert_soll).Value = Range("NeueCharge_CuZahlerwert").Value
    .DataBodyRange(1, QSilberDB_Col.AgWert_soll).Value = Range("NeueCharge_AgZahlerwert").Value
    .DataBodyRange(1, QSilberDB_Col.AgStrom_soll).Value = Range("NeueCharge_AgSollstrom").Value
  End With
  
  ' Print the Chargeschein
  Sheet7.Range("Zettel_ChargeNummer").Value = chargeNummer
  'PrintChargeschein

  ' Save and close the database
  wbook.SaveAs GDatabasePath
  app.Quit
  Set app = Nothing

  ' Refresh all data connections & queries
  ThisWorkbook.RefreshAll

  ' Show confirmation message
  Application.StatusBar = "Neue Charge für Material " & Range("E7").Value & " erfolgreich gespeichert"
  Notify Range("E7").Value, "Charge erfolgreich hinzugefügt", "Info", 10

  ' Reset the form
  FormReset True
End Sub

Private Sub NeueCharge_AbbrechenBtn_Click()
  FormReset True
End Sub

Function IsDatabaseOpen()
  Dim ff As Long, ErrNo As Long

  On Error Resume Next
  ff = FreeFile()
  Open GDatabasePath For Input Lock Read As #ff
  Close ff
  ErrNo = Err
  On Error GoTo 0

  Select Case ErrNo
    Case 0:    IsDatabaseOpen = False
    Case 70:   IsDatabaseOpen = True
    Case Else: Error ErrNo
  End Select
End Function

Private Sub PrintChargeschein()
  Dim chargePaper As Worksheet
  Set chargePaper = Sheet7 ' Zettel

  ' Set up the page
  With chargePaper.PageSetup
    .PaperSize = 70 ' Found with macro recorder, should be A6
    .Orientation = xlPortrait
    .FitToPagesWide = 1
    .FitToPagesTall = False
    .Zoom = False
    .BlackAndWhite = False
    .HeaderMargin = 0
    .FooterMargin = 0
    .TopMargin = 0.5
    .BottomMargin = 0.5
    .RightMargin = 0.5
    .LeftMargin = 0.5
    .CenterHorizontally = True
    .CenterVertically = True
  End With

  ' Print the page
  'chargePaper.PrintPreview
  chargePaper.PrintOut
End Sub
