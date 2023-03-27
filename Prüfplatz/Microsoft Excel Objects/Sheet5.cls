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
  Alterung_SaveBtn.Enabled = False
  Alterung_SaveBtn.BackColor = vbButtonFace
  If Range("IsFormValid").Value Then
    Alterung_SaveBtn.Enabled = True
    Alterung_SaveBtn.BackColor = RGB(0, 192, 0)
  End If
End Sub

Private Sub CreateMode_Fill()
' Compile the form in creation mode. Copies the values from the same
' Auftrag, if these values are found. That speeds up the filling of
' the form from the Employee.
' Warning: the values from database are retrieved by formulas in the
' Excel sheet and not by the macro.

  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword

  ' Suggest Löttest
  Range("Result").Value = "JA"

  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub EditMode_Fill()
' Compile the form in edit mode. Copies the values from the Quality
' Database to be modified by the Worker and then overwritten.
' Warning: values from database are retrieved by formulas in the Excel
' sheet and not by the macro. They are referenced as normal range
' pointing to the cells.

  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword

  ' Copy Alterung iO & niO
  If CStr(Range("ResultDB").Value) <> "" Then
    Range("Result").Value = Range("ResultDB").Value
  End If

  ' Copy worker number
  If CStr(Range("MitarbeiterDB").Value) <> "" Then
    Range("Mitarbeiter").Value = Range("MitarbeiterDB").Value
  End If

  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells  
End Sub

Private Sub FormReset(ByVal FullReset As Boolean)
' Resets the form, deleting all the custom data and moving the selection
' to the barcode Cell

  Set sheet = Sheet5

  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword

  ' Reset form
  If FullReset Then
    Range("ChargenummerCell").ClearContents
  End If
  Range("Result").ClearContents
  Range("Mitarbeiter").ClearContents

  Range("ChargenummerCell").Select
  ActiveWindow.Scrollrow = 1
  ActiveWindow.ScrollColumn = 1
  
  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub Alterung_DeleteBtn_Click()
  FormReset True

  ' Go to main scanning page
  Sheet2.Activate
End Sub

Private Sub Alterung_SaveBtn_Click()
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
  Application.StatusBar = "Datenbank wird geöffnet, bitte warten ..."

  ' Definition of variables
  Dim app As New Excel.Application
  Dim wbook As Excel.Workbook
  Dim sheet As Worksheet
  Dim table As ListObject
  Dim charges As Variant
  Dim i, j As Long
  Dim chargeRow As Single
  Dim databasePath As String
  databasePath = Sheet10.Range("ProjectPfad").Value & "Q_Silber_DATABASE.xlsx"

  ' Check if database is open
  For j = 1 To 30
    If IsDatabaseOpen(databasePath) Then
      Application.StatusBar = "Datenbank beschäftigt, bitte warten ..."
      If j = 30 Then
        Notify "FEHLER", "Charge wurde NICHT gespeichert, weil die Datenbank ausgelastet ist! Bitte ernuet probieren", Sheet10.Range("ProjectPfad").Value, "Error", 20
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
    .DataBodyRange(chargeRow, QSilberDB_Col.Alterung_iO).Value = Range("Result").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.Alterung_Mitarbeiter).Value = Range("Mitarbeiter").Value
    .DataBodyRange(chargerow, QSilberDB_Col.Alterung_Datum).Value = Range("Datum").Value
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
  FormReset True

  ' Go to main scanning page
  Sheet2.Activate
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

  If Range("IsAlterungNeeded").Value Then
    Dim chargeNummer
    chargeNummer = Range("ChargenummerValue").Value

    Select Case Range("CurrentChargePhase").Value
      Case chargePhase.Processing ' Process and quality data are missing
        MsgBox "Diese Charge befindet sich noch in Bearbeitung, bitte zuerst die Prozess- und Qualitätsdaten hinzufügen!"
        Sheet1.Activate
        Sheet1.Range("Prozess_Chargenummer").Value = chargeNummer
        Exit Function
      Case chargePhase.Nacharbeit ' Nacharbeit done and löttest needed
        ' TO DO
        ' Need to split workflow between edit and create?
      Case chargePhase.Alterung
        FormReset False
        CreateMode_Fill
      Case chargePhase.Done To chargePhase.Scrapped ' Charge is done or scrapped
        FormReset False
        EditMode_Fill
      Case Else
        MsgBox "Unspecified error with this charge, Phase not handled!"
        FormReset True
    End Select
  Else
    MsgBox "Dieses Teilnummer benötigt keine Löttest!"
    FormReset True
  End If
End Function