Option Explicit

Public sheet As Worksheet
Private silberBedarfCell, silberSchichtCell, anlageCell, barCodeCell, bezeichnungCell As Range
Private kupferSchichtCell, alterungCell, grundMaterialCell, nettoGewichtCell As Range
Private barCodeValueCell, isTeileInDBCell, isFormValidCell, isBarcodeValidCell As Range
Private CuEZCell, AgEZCell, AgStrom, isBarcodeEmptyCell As Range
Private AgStrasseEZCell, AgGlockeEZCell, AgStrasseStromCell, AgGlockeStromCell As Range

Private Sub Worksheet_Activate()
  FormReset True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
' Is triggered everytime a cell content changes in the Worksheet

  ' Define cell ranges
  Set barCodeCell = Range("NeueTeile_Barcode")
  Set isTeileInDBCell = Range("NeueTeile_IsTeileInDatabase")
  Set isFormValidCell = Range("NeueTeile_IsFormValid")
  Set isBarcodeValidCell = Range("NeueTeile_IsBarcodeValid")
  Set isBarcodeEmptyCell = Range("NeueTeile_IsBarcodeEmpty")
  
  ' Control if barcode cell was changed and act accordingly
  If Not Intersect(Target, barCodeCell) Is Nothing Then
    If Not isBarcodeEmptyCell.Value Then
      If isBarcodeValidCell.Value Then
        If isTeileInDBCell Then
          FormReset False
          EditMode_Fill
        Else
          FormReset False
          CreateMode_Fill
        End If
      Else
        MsgBox "Ungültiger barcode!"
        FormReset True
      End If
    End If
  End If
  
  ' Enable or disable the save button
  NeueTeile_SpeichernBtn.Enabled = False
  NeueTeile_SpeichernBtn.BackColor = vbButtonFace
  If isFormValidCell.Value Then
    NeueTeile_SpeichernBtn.Enabled = True
    NeueTeile_SpeichernBtn.BackColor = RGB(0, 192, 0)
  End If
  
  ' Enable or disable the delete button
  NeueTeile_LoeschenBtn.Enabled = False
  NeueTeile_LoeschenBtn.BackColor = vbButtonFace
  If isTeileInDBCell.Value Then
    NeueTeile_LoeschenBtn.Enabled = True
    NeueTeile_LoeschenBtn.BackColor = RGB(192, 0, 0)
  End If
End Sub

Private Sub FormReset(ByVal BarcodeReset As Boolean)
' Resets the form, deleting all the custom data and moving the selection
' to the barcode cell
  
  Set sheet = Sheet4
  
  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword
  
  ' Reset form
  If BarcodeReset Then
    Set barCodeCell = Range("NeueTeile_Barcode")
    barCodeCell.ClearContents
  End If
  
  Set bezeichnungCell = Range("NeueTeile_Bezeichnung")
  bezeichnungCell.ClearContents
  
  Set grundMaterialCell = Range("NeueTeile_Material")
  grundMaterialCell.ClearContents
  
  Set anlageCell = Range("NeueTeile_Anlage")
  anlageCell.ClearContents
  
  Set kupferSchichtCell = Range("NeueTeile_CuSchicht")
  kupferSchichtCell.ClearContents
  
  Set silberSchichtCell = Range("NeueTeile_AgSchicht")
  silberSchichtCell.ClearContents
  
  Set silberBedarfCell = Range("NeueTeile_AgBedarf")
  silberBedarfCell.ClearContents
  
  Set nettoGewichtCell = Range("NeueTeile_NettoGewicht")
  nettoGewichtCell.ClearContents
  
  Set alterungCell = Range("NeueTeile_Loettest")
  alterungCell.ClearContents
  
  Set CuEZCell = Range("NeueTeile_CuEZ")
  CuEZCell.ClearContents
  
  Set AgStrasseEZCell = Range("NeueTeile_AgEZStraße")
  AgStrasseEZCell.ClearContents
  Set AgStrasseStromCell = Range("NeueTeile_AgStromStraße")
  AgStrasseStromCell.ClearContents
  
  Set AgGlockeEZCell = Range("NeueTeile_AgEZGlocke")
  AgGlockeEZCell.ClearContents
  Set AgGlockeStromCell = Range("NeueTeile_AgStromGlocke")
  AgGlockeStromCell.ClearContents
  
  barCodeCell.Select
  
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
  Set bezeichnungCell = Range("NeueTeile_Bezeichnung")
  bezeichnungCell.Value = Range("I8").Value
  
  ' Copy "Grundmaterial"
  Set grundMaterialCell = Range("NeueTeile_Material")
  grundMaterialCell.Value = Range("I9").Value
  
  ' Copy "Standardanlage"
  Set anlageCell = Range("NeueTeile_Anlage")
  anlageCell.Value = Range("I11").Value
  
  ' Copy "Kupferschicht"
  Set kupferSchichtCell = Range("NeueTeile_CuSchicht")
  If CSng(Range("I13").Value) <> 0 Then
    kupferSchichtCell.Value = Range("I13").Value
  End If
  
  ' Copy "Silberschicht"
  Set silberSchichtCell = Range("NeueTeile_AgSchicht")
  If CSng(Range("I14").Value) <> 0 Then
    silberSchichtCell.Value = Range("I14").Value
  End If
  
  ' Copy "Silberbedarf"
  Set silberBedarfCell = Range("NeueTeile_AgBedarf")
  silberBedarfCell.Value = Range("I15").Value
  
  ' Copy "Gewicht/1000"
  Set nettoGewichtCell = Range("NeueTeile_NettoGewicht")
  nettoGewichtCell.Value = Range("I17").Value
  
  ' Copy "Alterung & Löttest"
  Set alterungCell = Range("NeueTeile_Loettest")
  If Range("I19").Value <> "" Then
    If Range("I19").Value Then
      alterungCell.Value = "JA"
    Else
      alterungCell.Value = "NEIN"
    End If
  End If
  
  ' Copy Cu Data from DB
  Set CuEZCell = Range("NeueTeile_CuEZ")
  If CSng(Range("I21").Value) > 0 Then
    CuEZCell.Value = Range("I21").Value
  End If
  
  ' Copy Ag Data from DB
  Set AgStrasseEZCell = Range("NeueTeile_AgEZStraße")
  If CSng(Range("I24").Value) > 0 Then
    AgStrasseEZCell.Value = Range("I24").Value
  End If
  Set AgStrasseStromCell = Range("NeueTeile_AgStromStraße")
  If CSng(Range("I25").Value) > 0 Then
    AgStrasseStromCell.Value = Range("I25").Value
  End If
  
  Set AgGlockeEZCell = Range("NeueTeile_AgEZGlocke")
  If CSng(Range("I26").Value) > 0 Then
    AgGlockeEZCell.Value = Range("I26").Value
  End If
  Set AgGlockeStromCell = Range("NeueTeile_AgStromGlocke")
  If CSng(Range("I27").Value) > 0 Then
    AgGlockeStromCell.Value = Range("I27").Value
  End If
End Sub

Private Sub CreateMode_Fill()
' Compile the form in creation mode. Copies only EZ values from the old EZ database
' so that the workers can fill the rest and everything will be saved in the DB.
' Warning: values from EZ database are retrieved by formulas in the Excel sheet and
' not by the macro. They are referenced as normal range pointing to the cells.
  
  ' Copy Cu Data from EZ DB
  Set CuEZCell = Range("NeueTeile_CuEZ")
  If CSng(Range("K21").Value) > 0 Then
    CuEZCell.Value = Range("K21").Value
  End If
  
  ' Copy Ag Data from DB
  Set AgStrasseEZCell = Range("NeueTeile_AgEZStraße")
  If CSng(Range("K24").Value) > 0 Then
    AgStrasseEZCell.Value = Range("K24").Value
  End If
  
  Set AgGlockeEZCell = Range("NeueTeile_AgEZGlocke")
  If CSng(Range("K26").Value) > 0 Then
    AgGlockeEZCell.Value = Range("K26").Value
  End If
End Sub

Private Sub NeueTeile_SpeichernBtn_Click()
' Save workflow triggered by Save button.
' If the Material is already in the database, it will call the Edit function
' If the Material is not in the database, it will call the Create function
'
' TASKS
' [x] Check if the Database file is already open
' [x] Wait for the Database file to close before saving
' [x] Update all the linked Query in the file after saving
' [x] Show saving message
' [x] Show better saved message
' [x] Split workflow between edit and save
' [x] Implement edit mode
' [x] Fuse save and edit together, workflow has to split only on the saving part!
' [ ] Add Kommentarfield

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
  
  ' Preparing data for write
  alterungDB = False
  If Range("NeueTeile_Loettest") = "JA" Then
    alterungDB = True
  End If
  
  teileData = Array( _
    Range("I7").Value, _
    Range("NeueTeile_Bezeichnung").Value, _
    Range("NeueTeile_Material").Value, _
    Range("NeueTeile_NettoGewicht").Value, _
    Range("NeueTeile_Anlage").Value, _
    Range("NeueTeile_CuSchicht").Value, _
    "", _
    Range("NeueTeile_CuEZ").Value, _
    "", _
    Range("NeueTeile_AgSchicht").Value, _
    Range("NeueTeile_AgBedarf").Value, _
    Range("NeueTeile_AgEZStraße").Value, _
    Range("NeueTeile_AgStromStraße").Value, _
    Range("NeueTeile_AgEZGlocke").Value, _
    Range("NeueTeile_AgStromGlocke").Value, _
    alterungDB _
  )
  ' I7 is Materialnummer
  ' Position 7 is Cu-Bedarf, leaved empty for future implementations
  ' Position 9 is Cu-Strom, leaved empty for future implementations
  ' Position 16 is for Löttest, we validate in the code
  
  ' Check if database is open
  Do While IsDatabaseOpen = True
    Application.StatusBar = "Datenbank beschäftigt, bitte warten ..."
    Application.Wait (Now + TimeValue("0:00:02"))
  Loop
  
  ' Opening of database file
  app.Visible = False ' Visible is False by default, so this should not be necessary
  app.DisplayAlerts = False
  Set wbook = app.Workbooks.Add(GDatabasePath)
  Application.StatusBar = "Wird gespeichert ..."
  
  ' Setting of targets
  Set sheet = wbook.Worksheets(2) ' Teil_DB
  Set table = sheet.ListObjects.Item(GTeilDB_TableName)
  
  ' Splitting workflow between Edit and Create
  isEditMode = Range("NeueTeile_IsTeileInDatabase").Value
  
  If isEditMode Then ' EDIT MODE
    ' Searching the row to overwrite
    '   Note: it is much faster to read a range of values into an array, then do calculations
    '   on that array, instead of doing the same calculation of the range of values itself
    '   I have found no way to only read a column, we need to read all the table
    materials = table.DataBodyRange
    For i = LBound(materials) To UBound(materials)
      If CStr(materials(i, 1)) = CStr(Range("I7").Value) Then
        teileRow = i
        Exit For
      End If
    Next i
  
    ' Write data in database
    For j = LBound(teileData) To UBound(teileData)
      table.DataBodyRange(teileRow, (j + 1)).Value = teileData(j)
    Next j
  Else ' CREATE MODE
    ' Write data in database
    table.ListRows.Add Position:=1
    For j = LBound(teileData) To UBound(teileData)
      table.DataBodyRange(1, (j + 1)).Value = teileData(j)
    Next j
  End If
  
  ' Save and close the database
  wbook.SaveAs GDatabasePath
  app.Quit
  Set app = Nothing
  
  ' Refresh all data connections & queries
  ThisWorkbook.RefreshAll
  
  ' Show confirmation message
  If isEditMode Then ' EDIT MODE
    Application.StatusBar = "Material-Nr. " & Range("E7").Value & " wurde aktualisiert"
    Notify Range("E7").Value, "Teile erfolghreich aktualisiert", "Info", 10
  Else ' CREATE MODE
    Application.StatusBar = "Material-Nr. " & Range("E7").Value & " wurde erstellt"
    Notify Range("E7").Value, "Teile erfolghreich hinzugefügt", "Info", 10
  End If
  
  ' Reset the form
  FormReset True
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
  Dim i As Long
  Dim teileRow As Integer

  If MsgBox("Möchten Sie dieses Material wirklich dauerhaft löschen?", vbCritical + vbYesNo, "Achtung") = vbYes Then

    ' Show saving message
    Application.StatusBar = "Datenbank wird geöffnet, bitte warten ..."

    ' Check if database is open
    Do While IsDatabaseOpen = True
      Application.StatusBar = "Datenbank beschäftigt, bitte warten ..."
      Application.Wait (Now + TimeValue("0:00:02"))
    Loop

    ' Opening of database file
    app.Visible = False ' Visible is False by default, so this should not be necessarz
    app.DisplayAlerts = False
    Set wbook = app.Workbooks.Add(GDatabasePath)
    Application.StatusBar = "Wird gespeichert ..."
    
    ' Setting of targets
    Set sheet = wbook.Worksheets(2) ' Teil_DB
    Set table = sheet.ListObjects.Item(GTeilDB_TableName)

    ' Searching the row to delete
    materials = table.DataBodyRange
    For i = LBound(materials) To UBound(materials)
      If CStr(materials(i,1)) = CStr(Range("I7").Value) Then
        teileRow = i
        Exit For
      End If
    Next i

    ' Delete the row
    table.ListRows(teileRow).Delete

    ' Save and close the database
    wbook.SaveAs GDatabasePath
    app.Quit
    Set app = Nothing

    ' Show confirmation message
    Application.StatusBar = "Material-Nr. " & Range("E7").Value & " wurde gelöscht"
    Notify Range("E7").Value, "Teile erfolgreich gelöscht", "Info", 10

    ' Reset the form
    FormReset True

    ' Refresh all data connections & queries
    ThisWorkbook.RefreshAll
  End If
End Sub

Private Sub NeueTeile_ResetBtn_Click()
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
