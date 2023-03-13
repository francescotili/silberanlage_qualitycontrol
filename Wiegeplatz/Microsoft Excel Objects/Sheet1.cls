Option Explicit

Public sheet As Worksheet

Private Sub Worksheet_Activate()
  FormReset True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
' Is triggered everytime a cell content changes in the Worksheet
  
  ' Control if barcode cell was changed and act accordingly
  If Not Intersect(Target, Range("BarcodeCell")) Is Nothing Then
    If Not Range("IsBarcodeEmpty").Value Then
      If Range("IsBarcodeValid").Value Then
        If Range("IsTeileInDatabase").Value Then
          FormReset False
          Form_Fill
        Else
          If MsgBox("Teilenummer nicht vorhanden!" & vbNewLine & "Jetzt hinzufügen?", vbInformation + vbYesNo + vbDefaultButton1, "ACHTUNG") = vbYes Then
            Sheet4.Activate
            Sheet4.Range("NeueTeile_Barcode").Value = Range("BarcodeValue").Value
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
  If Range("IsFormValid").Value Then
    NeueCharge_SpeichernBtn.Enabled = True
    NeueCharge_SpeichernBtn.BackColor = RGB(0, 192, 0)
  End If
End Sub

Private Sub FormReset(ByVal FullReset As Boolean)
' Resets the form, deleting all the custom data and moving the selection
' to the barcode cell

  Set sheet = Sheet1
  
  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword
  
  ' Reset form
  If FullReset Then
    Range("BarcodeCell").ClearContents
  End If
  Range("Gewicht").ClearContents
  Range("Anlage").ClearContents
  'Range("Mitarbeiter").ClearContents
  Range("Kommentar").ClearContents
  
  Range("BarcodeCell").Select
  
  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub Form_Fill()
' Compile the form with the data scanned from the Kanban Karte and the data
' retrieved from the database
  
  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword
  
  ' Copy "Standardanlage"
  Range("Anlage").Value = Range("AnlageSoll").Value
  
  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub NeueCharge_SpeichernBtn_Click()
' Save workflow triggered by Save button.
' This form will always create a new Chargenummer and add it to the
' Production database
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
  chargeNummer = Range("Chargenummer").Value

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
    .DataBodyRange(1, QSilberDB_Col.chargeNummer).Value = chargeNummer
    .DataBodyRange(1, QSilberDB_Col.Datum).Value = Range("Datum").Value
    .DataBodyRange(1, QSilberDB_Col.Annahme_Mitarbeiter).Value = Range("Mitarbeiter").Value
    .DataBodyRange(1, QSilberDB_Col.Annhame_Kommentar).Value = Range("Kommentar").Value
    .DataBodyRange(1, QSilberDB_Col.Materialnummer).Value = Range("Materialnummer").Value
    .DataBodyRange(1, QSilberDB_Col.Gewicht_netto).Value = Range("NettoGewicht").Value
    .DataBodyRange(1, QSilberDB_Col.CuSchicht_soll).Value = Range("CuSchichtSoll").Value
    .DataBodyRange(1, QSilberDB_Col.AgSchicht_soll).Value = Range("AgSchichtSoll").Value
    .DataBodyRange(1, QSilberDB_Col.Auftragsnummer).Value = Range("Auftragsnummer").Value
    .DataBodyRange(1, QSilberDB_Col.Füllmenge).Value = Range("FüllmengeAFO").Value
    .DataBodyRange(1, QSilberDB_Col.ChargeGewicht).Value = Range("Gewicht").Value
    .DataBodyRange(1, QSilberDB_Col.Stückzahl).Value = Range("Stückzahl").Value
    .DataBodyRange(1, QSilberDB_Col.AgBedarf_soll).Value = Range("AgBedarfSollTOT").Value
    .DataBodyRange(1, QSilberDB_Col.Anlage).Value = Range("Anlage").Value
    .DataBodyRange(1, QSilberDB_Col.CuWert_soll).Value = Range("CuZahlerWert").Value
    .DataBodyRange(1, QSilberDB_Col.AgWert_soll).Value = Range("AgZahlerWert").Value
    .DataBodyRange(1, QSilberDB_Col.AgStrom_soll).Value = Range("AgStromSoll").Value
    .DataBodyRange(1, QSilberDB_Col.Phase).Value = ChargePhase.Processing
  End With
  
  ' Print the Chargeschein
  Sheet7.Range("Chargenummer").Value = chargeNummer
  PrintChargeschein

  ' Save and close the database
  wbook.SaveAs GDatabasePath
  app.Quit
  Set app = Nothing

  ' Refresh all data connections & queries
  ThisWorkbook.RefreshAll

  ' Show confirmation message
  Application.StatusBar = "Neue Charge für Material " & Range("MaterialnummerString").Value & " erfolgreich gespeichert"
  Notify Range("MaterialnummerString").Value, "Charge erfolgreich hinzugefügt", "Info", 10

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
  ThisWorkbook.Unprotect Password:=GAdminPassword
  chargePaper.Visible = True ' Sheet needs to be visible to be printed
  chargePaper.PrintPreview
  'chargePaper.PrintOut
  chargePaper.Visible = xlSheetHidden
  ThisWorkbook.Protect Password:=GAdminPassword
End Sub
