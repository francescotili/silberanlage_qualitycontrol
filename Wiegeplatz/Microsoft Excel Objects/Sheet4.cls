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
          EditWorkflow
        Else
          ThisWorkbook.RefreshAll
          DoEvents ' Waits for the refresh to finish
          If Range("IsTeileInDatabase").Value Then
            EditWorkflow
          Else
            FormReset False
            CreateMode_Fill
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
  NeueTeile_SpeichernBtn.Enabled = False
  NeueTeile_SpeichernBtn.BackColor = vbButtonFace
  If Range("IsFormValid").Value Then
    NeueTeile_SpeichernBtn.Enabled = True
    NeueTeile_SpeichernBtn.BackColor = RGB(0, 192, 0)
  End If
  
  ' Enable or disable the delete button
  NeueTeile_LoeschenBtn.Enabled = False
  NeueTeile_LoeschenBtn.BackColor = vbButtonFace
  If Range("IsTeileInDatabase").Value Then
    NeueTeile_LoeschenBtn.Enabled = True
    NeueTeile_LoeschenBtn.BackColor = RGB(192, 0, 0)
  End If
End Sub

Private Sub FormReset(ByVal FullReset As Boolean)
' Resets the form, deleting all the custom data and moving the selection
' to the barcode cell
  
  Set sheet = Sheet4
  
  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword
  
  ' Reset form
  If FullReset Then
    Range("BarcodeCell").ClearContents
  End If
  Range("Bezeichnung").ClearContents
  Range("Grundmaterial").ClearContents
  Range("AnlageSoll").ClearContents
  Range("CuSchicht").ClearContents
  Range("AgSchicht").ClearContents
  Range("AgBedarf").ClearContents
  Range("NettoGewicht").ClearContents
  Range("IsAlterungNeeded").ClearContents
  Range("CuEZ").ClearContents
  Range("AgEZ_Straße").ClearContents
  Range("AgStrom_Straße").ClearContents
  Range("AgEZ_Glocke").ClearContents
  Range("AgStrom_Glocke").ClearContents
  Range("Mitarbeiter").ClearContents
  
  Range("BarcodeCell").Select
  ActiveWindow.Scrollrow = 1
  ActiveWindow.ScrollColumn = 1
  
  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub EditMode_Fill()
' Compile the form in edit mode. Copies the values from the database
' to be modified by the Worker and then saved.
' Warning: values from database are retrieved by formulas in the Excel sheet and
' not by the macro. They are referenced as normal range pointing to the cells.
  
  ' Copy "Bezeichnung"
  If CStr(Range("BezeichnungDB").Value) <> "" Then
    Range("Bezeichnung").Value = Range("BezeichnungDB").Value
  End If
  
  ' Copy "Grundmaterial"
  If CStr(Range("GrundmaterialDB").Value) <> "" Then
    Range("Grundmaterial").Value = Range("GrundmaterialDB").Value
  End If
  
  ' Copy "Standardanlage"
  If CStr(Range("AnlageSollDB").Value) <> "" Then
    Range("AnlageSoll").Value = Range("AnlageSollDB").Value
  End If
  
  ' Copy "Kupferschicht"
  If CSng(Range("CuSchichtDB").Value) <> 0 Then
    Range("CuSchicht").Value = Range("CuSchichtDB").Value
  End If
  
  ' Copy "Silberschicht"
  If CSng(Range("AgSchichtDB").Value) <> 0 Then
    Range("AgSchicht").Value = Range("AgSchichtDB").Value
  End If
  
  ' Copy "Silberbedarf"
  If CSng(Range("AgBedarfDB").Value) <> 0 Then
    Range("AgBedarf").Value = Range("AgBedarfDB").Value
  End If
  
  ' Copy "Gewicht/1000"
  If CSng(Range("NettoGewichtDB").Value) <> 0 Then
    Range("NettoGewicht").Value = Range("NettoGewichtDB").Value
  End If
  
  ' Copy "Alterung & Löttest"
  If Range("IsAlterungNeededDB").Value <> "" Then
    If Range("IsAlterungNeededDB").Value Then
      Range("IsAlterungNeeded").Value = "JA"
    Else
      Range("IsAlterungNeeded").Value = "NEIN"
    End If
  End If
  
  ' Copy Cu Data from DB
  If CSng(Range("CuEZDB").Value) > 0 Then
    Range("CuEZ").Value = Range("CuEZDB").Value
  End If
  
  ' Copy Ag Data from DB for Straße
  If CSng(Range("AgEZDB_Straße").Value) > 0 Then
    Range("AgEZ_Straße").Value = Range("AgEZDB_Straße").Value
  End If
  If CSng(Range("AgStromDB_Straße").Value) > 0 Then
    Range("AgStrom_Straße").Value = Range("AgStromDB_Straße").Value
  End If
  
  ' Copy Ag Data from DB for Glocke
  If CSng(Range("AgEZDB_Glocke").Value) > 0 Then
    Range("AgEZ_Glocke").Value = Range("AgEZDB_Glocke").Value
  End If
  If CSng(Range("AgStromDB_Glocke").Value) > 0 Then
    Range("AgStrom_Glocke").Value = Range("AgStromDB_Glocke").Value
  End If
End Sub

Private Sub CreateMode_Fill()
' Compile the form in creation mode. Copies only EZ values from the old EZ database
' so that the workers can fill the rest and everything will be saved in the DB.
' Warning: values from EZ database are retrieved by formulas in the Excel sheet and
' not by the macro. They are referenced as normal range pointing to the cells.
  
  ' Copy Cu Data from Old EZ DB
  If CStr(Range("CuOldEZ").Value) <> "" Then
    If CSng(Range("CuOldEZ").Value) > 0 Then
      Range("CuEZ").Value = Range("CuOldEZ").Value
    End If
  End If
  
  ' Copy Ag Data from Old DB for Straße
  If CStr(Range("AgOldEZ_Straße").Value) <> "" Then
    If CSng(Range("AgOldEZ_Straße").Value) > 0 Then
      Range("AgEZ_Straße").Value = Range("AgOldEZ_Straße").Value
    End If
  End If
  
  ' Copy Ag Data from Old DB for Glocke
  If CStr(Range("AgOldEZ_Glocke").Value) <> "" Then
    If CSng(Range("AgOldEZ_Glocke").Value) > 0 Then
      Range("AgEZ_Glocke").Value = Range("AgOldEZ_Glocke").Value
    End If
  End If
End Sub

Private Sub NeueTeile_SpeichernBtn_Click()
' Save workflow triggered by Save button.
' If the Material is already in the database, it will call the Edit function
' If the Material is not in the database, it will call the Create function
'
' TASKS
' [ ] Add Kommentar as Dialog window

  SaveMaterial
End Sub

Private Sub SaveMaterial()
' Reads the form data and compile an array to write. Then opens the database file
' in hidden mode and write the data in the database. If the material is already there
' it will overwritten

  ' Show saving message
  Application.StatusBar = "Datenbank wird geöffnet, bitte warten ..."
  
  ' Definition of variables
  Dim app As New excel.Application
  Dim wbook As excel.Workbook
  Dim sheet As Worksheet
  Dim table As ListObject
  Dim materials As Variant
  Dim teileData As Variant
  Dim i, j As Long
  Dim newRow As ListRow
  Dim teileRow As Integer
  Dim alterungDB As Boolean
  Dim isEditMode As Boolean
  Dim databasePath As String
  databasePath = Sheet5.Range("ProjectPfad").Value & "Q_Silber_DATABASE.xlsx"
  
  ' Preparing data for write
  alterungDB = False
  If Range("IsAlterungNeeded") = "JA" Then
    alterungDB = True
  End If
  
  ' Check if database is open
  For j = 1 To 30
    If IsDatabaseOpen(databasePath) Then
      Application.StatusBar = "Datenbank beschäftigt, bitte warten ..."
      If j = 30 Then
        Notify "FEHLER", "Charge wurde NICHT gespeichert, weil die Datenbank ausgelastet ist! Bitte ernuet probieren", Sheet5.Range("ProjectPfad").Value, "Error", 20
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
  Set sheet = wbook.Worksheets("Teil_DB")
  Set table = sheet.ListObjects.Item("Teiledatenbank")
  
  ' Splitting workflow between Edit and Create
  If Range("IsTeileInDatabase").Value Then
    ' Search the row to overwrite
    materials = table.DataBodyRange
    For i = LBound(materials) To UBound(materials)
      If CStr(materials(i, 1)) = CStr(Range("Materialnummer").Value) Then
        teileRow = i
        Exit For
      End If
    Next i
  Else
    ' Create a new row
    table.ListRows.Add Position:=1
    teileRow = 1
  End If

  ' Write data in database
  With table
    .DataBodyRange(teileRow, TeilDB_Col.Materialnummer).Value = Range("Materialnummer").Value
    .DataBodyRange(teileRow, TeilDB_Col.Bezeichnung).Value = Range("Bezeichnung").Value
    .DataBodyRange(teileRow, TeilDB_Col.Grundmaterial).Value = Range("Grundmaterial").Value
    .DataBodyRange(teileRow, TeilDB_Col.Gewicht_netto).Value = Range("NettoGewicht").Value
    .DataBodyRange(teileRow, TeilDB_Col.Anlage_soll).Value = Range("AnlageSoll").Value
    .DataBodyRange(teileRow, TeilDB_Col.CuSchicht).Value = Range("CuSchicht").Value
    .DataBodyRange(teileRow, TeilDB_Col.CuEZ).Value = Range("CuEZ").Value
    .DataBodyRange(teileRow, TeilDB_Col.AgSchicht).Value = Range("AgSchicht").Value
    .DataBodyRange(teileRow, TeilDB_Col.AgBedarf).Value = Range("AgBedarf").Value
    .DataBodyRange(teileRow, TeilDB_Col.AgEZ_Straße).Value = Range("AgEZ_Straße").Value
    .DataBodyRange(teileRow, TeilDB_Col.AgStrom_Straße).Value = Range("AgStrom_Straße").Value
    .DataBodyRange(teileRow, TeilDB_Col.AgEZ_Glocke).Value = Range("AgEZ_Glocke").Value
    .DataBodyRange(teileRow, TeilDB_Col.AgStrom_Glocke).Value = Range("AgStrom_Glocke").Value
    .DataBodyRange(teileRow, TeilDB_Col.Löten).Value = alterungDB
  End With
  If Range("IsTeileInDatabase").Value Then
    table.DataBodyRange(teileRow, TeilDB_Col.Editor).Value = Range("Mitarbeiter").Value
  Else
    table.DataBodyRange(teileRow, TeilDB_Col.Ersteller).Value = Range("Mitarbeiter").Value
  End If
  
  ' Save and close the database
  wbook.SaveAs databasePath
  app.Quit
  Set app = Nothing
  
  ' Refresh all data connections & queries
  ThisWorkbook.RefreshAll
  
  ' Show confirmation message
  If Range("IsTeileInDatabase").Value Then ' EDIT MODE
    Application.StatusBar = "Material-Nr. " & Range("MaterialnummerString").Value & " wurde aktualisiert"
    Notify Range("MaterialnummerString").Value, "Teile erfolgreich aktualisiert", Sheet5.Range("ProjectPfad").Value, "Info", 10
  Else ' CREATE MODE
    Application.StatusBar = "Material-Nr. " & Range("MaterialnummerString").Value & " wurde erstellt"
    Notify Range("MaterialnummerString").Value, "Teile erfolgreich hinzugefügt", Sheet5.Range("ProjectPfad").Value, "Info", 10
  End If
  
  ' Reset the form
  FormReset True

  ' Go back to "Neue Charge" worksheet
  Sheet1.Activate
End Sub

Private Sub NeueTeile_LoeschenBtn_Click()
' Open the database file in hidden mode and delete the specified Teilnummer.
' It will delete a row only if it can find it

  ' Definition of variables
  Dim app As New excel.Application
  Dim wbook As excel.Workbook
  Dim sheet As Worksheet
  Dim table As ListObject
  Dim materials As Variant
  Dim i, j As Long
  Dim teileRow As Integer
  Dim databasePath As String
  databasePath = Sheet5.Range("ProjectPfad").Value & "Q_Silber_DATABASE.xlsx"

  If MsgBox("Möchten Sie dieses Material wirklich dauerhaft löschen?", vbCritical + vbYesNo, "Achtung") = vbYes Then

    ' Show saving message
    Application.StatusBar = "Datenbank wird geöffnet, bitte warten ..."

    ' Check if database is open
  For j = 1 To 30
    If IsDatabaseOpen(databasePath) Then
      Application.StatusBar = "Datenbank beschäftigt, bitte warten ..."
      If j = 30 Then
        Notify "FEHLER", "Charge wurde NICHT gespeichert, weil die Datenbank ausgelastet ist! Bitte ernuet probieren", Sheet5.Range("ProjectPfad").Value, "Error", 20
        Exit Sub
      Else
        Application.Wait (Now + TimeValue("0:00:02"))
      End If
    Else
      Exit For
    End If
  Next j

    ' Opening of database file
    app.Visible = False ' Visible is False by default, so this should not be necessarz
    app.DisplayAlerts = False
    Set wbook = app.Workbooks.Add(databasePath)
    Application.StatusBar = "Wird gespeichert ..."
    
    ' Setting of targets
    Set sheet = wbook.Worksheets(Sheet1.Index) ' Teil_DB
    Set table = sheet.ListObjects.Item(1)

    ' Searching the row to delete
    materials = table.DataBodyRange
    For i = LBound(materials) To UBound(materials)
      If CStr(materials(i, 1)) = CStr(Range("Materialnummer").Value) Then
        teileRow = i
        Exit For
      End If
    Next i

    ' Delete the row
    table.ListRows(teileRow).Delete

    ' Save and close the database
    wbook.SaveAs databasePath
    app.Quit
    Set app = Nothing

    ' Show confirmation message
    Application.StatusBar = "Material-Nr. " & Range("MaterialnummerString").Value & " wurde gelöscht"
    Notify Range("MaterialnummerString").Value, "Teile erfolgreich gelöscht", Sheet5.Range("ProjectPfad").Value, "Info", 10

    ' Reset the form
    FormReset True

    ' Refresh all data connections & queries
    ThisWorkbook.RefreshAll
  End If
End Sub

Private Sub NeueTeile_ResetBtn_Click()
  FormReset True
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

Function EditWorkflow()
' The Teil is in DB, decide what to do next
  FormReset False
  EditMode_Fill
End Function