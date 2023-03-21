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
            MsgBox "Die Charge wurde nicht in der Datenbank gefunden!"
            FormReset True
          End If
        End If
      Else
        MsgBox "Ungültiger barcode!"
        FormReset True
      End If
    Else
      FormReset False
    End If
  End If

  ' Enable or disable the save button
  Prozess_SaveBtn.Enabled = False
  Prozess_SaveBtn.BackColor = vbButtonFace
  If Range("IsFormValid").Value Then
    Prozess_SaveBtn.Enabled = True
    Prozess_SaveBtn.BackColor = RGB(0, 192, 0)
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
    Range("CuZahlerIst").ClearContents
    Range("AgZahlerIst").ClearContents
    Range("AgStromIst").ClearContents
    Range("ChargenummerCell").ClearContents
  End If
  Range("Trommel").ClearContents
  Range("Mitarbeiter").ClearContents
  Range("Kommentar").ClearContents
  Range("CuSchichtIst").ClearContents
  Range("AgSchichtIst").ClearContents
  Range("Biegetest").ClearContents

  Range("ChargenummerCell").Select

  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub EditMode_Fill()
' Compile the form in edit mode. Copies the values from the Quality
' Database to be modified by the Worker and then saved.
' Warning: values from database are retrieved by formulas in the Excel
' sheet and not by the macro. They are referenced as normal range
' pointing to the cells.

  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword

  ' Copy "Trommel"
  If CStr(Range("TrommelDB").Value) <> "" Then
    Range("Trommel").Value = CStr(Range("TrommelDB").Value)
  End If

  ' Copy "Kupferdaten"
  If CStr(Range("CuZahlerIstDB").Value) <> "" Then
    Range("CuZahlerIst").Value = CSng(Range("CuZahlerIstDB").Value)
  End If
  If CStr(Range("CuSchichtIstDB").Value) <> "" Then
    Range("CuSchichtIst").Value = CSng(Range("CuSchichtIstDB").Value)
  End If

  ' Copy "Silberdaten"
  If CStr(Range("AgZahlerIstDB").Value) <> "" Then
    Range("AgZahlerIst").Value = CSng(Range("AgZahlerIstDB").Value)
  End If
  If CStr(Range("AgStromIstDB").Value) <> "" Then
    Range("AgStromIst").Value = CSng(Range("AgStromIstDB").Value)
  End If
  If CStr(Range("AgSchichtIstDB").Value) <> "" Then
    Range("AgSchichtIst").Value = CSng(Range("AgSchichtIstDB").Value)
  End If

  ' Copy "Biegetest"
  If CStr(Range("BiegetestDB").Value) <> "" Then
    Range("Biegetest").Value = CStr(Range("BiegetestDB").Value)
  End If

  ' Copy Prüfer & Kommentar
  If CStr(Range("MitarbeiterDB").Value) <> "" Then
    Range("Mitarbeiter").Value = CStr(Range("MitarbeiterDB").Value)
  End If
  If CStr(Range("KommentarDB").Value) <> "" Then
    Range("Kommentar").Value = CStr(Range("KommentarDB").Value)
  End If

  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub Prozess_DeleteBtn_Click()
  FormReset True

  ' Go to main scanning page
  Sheet2.Activate
  Sheet2.Range("ChargenummerCell").Select
End Sub

Private Sub Prozess_SaveBtn_Click()
' Save workflow triggered by Save button.
' If the Charge is in creation mode, it has no data for process
' and testing. That means that the data will be filled in the database
' at the corresponding row, after row search.
' If the Charge is in edit mode, this data will be overwritten.

  SaveCharge

  ' Go to main scanning page
  Sheet2.Activate
  Sheet2.Range("ChargenummerCell").Select
End Sub

Private Sub SaveCharge()
' Reads the form data and compile an array to write. Then opens the database file
' in hidden mode, find the row and write the data in the database.

  ' Show saving message
  Application.StatusBar = "Datenbank wird geöffnet, bitte warten ..."

  ' Definition of variables
  Dim app As New Excel.Application
  Dim wbook As Excel.Workbook
  Dim sheet As Worksheet
  Dim table As ListObject
  Dim charges As Variant
  Dim i As Long
  Dim chargeRow As Single
  Dim databasePath As String
  databasePath = Sheet10.Range("ProjectPfad").Value & "Q_Silber_DATABASE.xlsx"

  ' Check if database is open
  Do While IsDatabaseOpen(databasePath) = True
    Application.StatusBar = "Datenbank beschäftigt, bitte warten ..."
    Application.Wait (Now + TimeValue("0:00:02"))
  Loop
  
  ' Opening of database file
  app.Visible = False ' Visible is False by default, so this should not be necessary
  app.DisplayAlerts = False
  Set wbook = app.Workbooks.Add(databasePath)
  Application.StatusBar = "Wird gespeichert ..."
  
  ' Setting of targets
  Set sheet = wbook.Worksheets("Database")
  Set table = sheet.ListObjects.Item("Qualitätsdatabase")

  ' Searching the row
  charges = table.DataBodyRange
  For i = LBound(charges) To UBound(charges)
    If CStr(charges(i, 1)) = CStr(Range("ChargenummerValue").Value) Then
      chargeRow = i
      Exit For
    End If
  Next i

  ' Write data in database
  With table
    .DataBodyRange(chargeRow, QSilberDB_Col.Schicht).Value = Range("Schicht").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.Trommel).Value = Range("Trommel").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.CuWert_ist).Value = Range("CuZahlerIst").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.CuSchicht_ist).Value = Range("CuSchichtIst").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.AgWert_ist).Value = Range("AgZahlerIst").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.AgStrom_ist).Value = Range("AgStromIst").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.AgSchicht_ist).Value = Range("AgSchichtIst").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.Biegetest_iO).Value = Range("Biegetest").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.Prozess_Mitarbeiter).Value = Range("Mitarbeiter").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.Prozess_Kommentar).Value = Range("Kommentar").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.Phase).Value = Range("NextChargePhase").Value
  End With

  ' Save and close the database
  wbook.SaveAs databasePath
  app.Quit
  Set app = Nothing

  ' Show confirmation message
  Application.StatusBar = "Charge für Auftrag " & Range("Auftragsnummer").Value & " erfolgreich ergänzt"
  Notify "Auftrag " & Range("Auftragsnummer").Value, "Charge erfolgreich ergänzt", Sheet10.Range("ProjectPfad").Value, "Info", 10

  ' Refresh all data connections & queries
  ThisWorkbook.RefreshAll

  ' Reset the form
  Range("ChargenummerCell").ClearContents
  FormReset False
End Sub

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

  If Range("CurrentChargePhase").Value = 0 Then
    FormReset False
  Else
    FormReset False
    EditMode_Fill
  End If
End Function