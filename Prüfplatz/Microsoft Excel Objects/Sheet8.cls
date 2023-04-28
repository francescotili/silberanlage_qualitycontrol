Option Explicit

Public sheet As Worksheet

Private Sub Worksheet_Activate()
  FormReset True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
' Is triggered everytime a cell content changes in the Worksheet

  ' Control if barcode cell was changed and act accordingly
  If Not Intersect(Target, Range("ChargenummerCell")) Is Nothing Then
    If Not Range("IsBarcodeEmpty").Value Then
      If Range("IsBarcodeValid").Value Then
        If Range("IsChargeInDB").Value Then
          ChargeWorkflow
        Else
          ThisWorkbook.RefreshAll
          DoEvents ' Waits for the refresh to finish
          If Range("IsChargeInDB").Value Then
            ChargeWorkflow
          Else
            MsgBox str_chargeNotFound
            FormReset True
          End If
        End If
      Else
        MsgBox str_invalidBarcode
        FormReset True
      End If
    Else
      FormReset False
    End If
  End If

  ' Enable or disable the save button
  SaveBtn.Enabled = False
  SaveBtn.BackColor = vbButtonFace
  If Range("IsFormValid").Value Then
    SaveBtn.Enabled = True
    SaveBtn.BackColor = RGB(0, 192, 0)
  End If
End Sub

Private Sub EditMode_Fill()
' Compile the form in edit mode. Copies the values from the Quality
' Database to be modified by the Worker and then overwritten.
' Warning: values from database are retrieved by formulas in the Excel
' sheet and not by the macro. They are referenced as normal range
' pointing to the cells.

  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword

  ' Copy Charge Gewicht
  If CStr(Range("GewichtDB").Value) <> "" Then
    Range("Gewicht").Value = Range("GewichtDB").Value
  End If

  ' Copy Anlage
  If CStr(Range("AnlageDB").Value) <> "" Then
    Range("Anlage").Value = Range("AnlageDB").Value
  End If

  ' Copy Mitarbeier
  If CStr(Range("MitarbeiterDB").Value) <> "" Then
    Range("Mitarbeiter").Value = Range("MitarbeiterDB").Value
  End If

  ' Copy Kommentar
  If CStr(Range("KommentarDB").Value) <> "" Then
    Range("Kommentar").Value = Range("KommentarDB").Value
  End If

  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells  
End Sub

Private Sub FormReset(ByVal FullReset As Boolean)
' Resets the form, deleting all the custom data and moving the selection
' to the barcode Cell
  Set sheet = Sheet8

  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword

  ' Reset form
  If FullReset Then
    Range("ChargenummerCell").ClearContents
  End If
  Range("Gewicht").ClearContents
  Range("Anlage").ClearContents
  Range("Mitarbeiter").ClearContents
  Range("Kommentar").ClearContents

  Range("ChargenummerCell").Select
  ActiveWindow.Scrollrow = 1
  ActiveWindow.ScrollColumn = 1
  
  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub AbbrechenBtn_Click()
  FormReset True

  ' Go to main scanning page
  Sheet2.Activate
End Sub

Private Sub SaveBtn_Click()
' Save workflow triggered by Save button.
' If the Charge is in creation mode, it has no data for process
' and testing. That means that the data will be filled in the database
' at the corresponding row, after row search.
' If the Charge is in edit mode, this data will be overwritten.

  SaveCharge
End Sub

Private Sub SaveCharge()
' Reads the form data and compile an array to write. Then opens the database file
' in hidden mode, find the row and write the data in the database.

  ' Show saving message
  Application.StatusBar = str_statusBar_DBOpening

  ' Definition of variables
  Dim app As New Excel.Application
  Dim wbook As Excel.Workbook
  Dim sheet As Worksheet
  Dim table As ListObject
  Dim charges As Variant
  Dim i, j As Long
  Dim chargeRow As Single
  Dim databasePath As String
  databasePath = Sheet10.Range("ProjectPfad").Value & "Q_Silber_DATABASE.xlsm"

  ' Check if database is open
  For j = 1 To 30
    If IsDatabaseOpen(databasePath) Then
      Application.StatusBar = str_statusBar_DBBusy
      If j = 30 Then
        Notify str_notify_errorTitle, str_notify_savingDatabaseBusy, Sheet10.Range("ProjectPfad").Value, "Error", 20
        Exit Sub
      Else
        Application.Wait (Now + TimeValue("0:00:02"))
      End If
    Else
      Exit For
    End If
  Next j
  
  ' Opening of database file
  app.Visible = False ' Visible is False by default, so this should not be necessary
  app.DisplayAlerts = False
  Set wbook = app.Workbooks.Add(databasePath)
  Application.StatusBar = str_statusBar_DBSaving
  
  ' Setting of targets
  Set sheet = wbook.Worksheets("Database")
  Set table = sheet.ListObjects.Item("Qualitätsdatabase")

  ' Searching the row
  charges = table.DataBodyRange
  For i = LBound(charges) To UBound(charges)
    If CStr(charges(i, QSilberDB_Col.ChrgNummer)) = CStr(Range("ChargenummerValue").Value) Then
      chargeRow = i
      Exit For
    End If
  Next i

  ' Write data in database
  With table
    .DataBodyRange(chargeRow, QSilberDB_Col.ChargeGewicht).Value = Range("Gewicht").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.Anlage).Value = Range("Anlage").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.Annahme_Mitarbeiter).Value = Range("Mitarbeiter").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.Annhame_Kommentar).Value = Range("Kommentar").Value
  End With

  ' Save and close the database
  wbook.SaveAs Filename:=databasePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
  app.Quit
  Set app = Nothing

  ' Show confirmation message
  Application.StatusBar = StringFormat(str_statusBar_chargeSaved, Range("Auftragsnummer").Value)
  Notify StringFormat(str_notify_chargeSavedTitle, Range("Auftragsnummer").Value), str_notify_chargeSavedText, Sheet10.Range("ProjectPfad").Value, "Info", 10

  ' Refresh all data connections & queries
  ThisWorkbook.RefreshAll

  ' Reset the form
  FormReset True

  ' Go to main scanning page
  Sheet2.Activate
End sub

Function IsDatabaseOpen(databasePath As String)
  Dim ff As Long, ErrNo As Long

  On Error Resume Next
  ff = FreeFile()
  Open databasePath For Input Lock Read As #ff
  Close ff
  ErrNo = Err
  On Error GoTo 0

  Select Case ErrNo
    Case 0:    IsDatabaseOpen = False
    Case 70:   IsDatabaseOpen = True
    Case Else: Error ErrNo
  End Select
End Function

Function ChargeWorkflow()
' The charge is in DB, decide what to do next

  ' There is no "create" mode, because if the Annahme has to be
  ' created, that means that the Chargenummer cannot be found
  ' So, it is always only Edit mode
  FormReset False
  EditMode_Fill
End Function