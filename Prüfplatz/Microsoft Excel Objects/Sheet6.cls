Option Explicit

Public sheet As Worksheet

Private Sub Worksheet_Activate()
  FormReset
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
' Is triggered everytime a cell content changes in the Worksheet

  ' Control if barcode cell was changed and act accordingly
  If Not Intersect(Target, Range("ChargenummerCell")) Is Nothing Then
    If Not Range("IsBarcodeEmpty").Value Then
      If Range("IsBarcodeValid").Value Then
        If Not Range("IsChargeInDB").Value Then ' If IsChargeInDB form is automatically filled
          ThisWorkbook.RefreshAll
          DoEvents ' Waits for the refresh to finish
          If Not Range("IsChargeInDB").Value Then
            MsgBox str_chargeNotFound
            FormReset
          End If
        End If
      Else
        MsgBox str_invalidBarcode
        FormReset
      End If
    End If
  End If

  ' Enable or disable the various buttons
  Annahme_Btn.Enabled = False
  Prozess_Btn.Enabled = False
  Alterung_Btn.Enabled = False
  Nacharbeit_Btn.Enabled = False
  Scrap_Btn.Enabled = False
  Delete_Btn.Enabled = False
  Delete_Btn.BackColor = vbButtonFace
  If Range("IsChargeInDB").Value Then
    Annahme_Btn.Enabled = True
    Prozess_Btn.Enabled = True
    If Range("IsAlterungNeeded").Value Then
      Alterung_Btn.Enabled = True
    End If
    Nacharbeit_Btn.Enabled = True
    Scrap_Btn.Enabled = True
    Delete_Btn.Enabled = True
    Delete_Btn.BackColor = RGB(192, 0, 0)
  End If
End Sub

Private Sub AbbrechenBtn_Click()
  FormReset

  ' Go to main scanning page
  Sheet2.Activate
End Sub

Private Sub Annahme_Btn_Click()
  Dim chargeNummer
  chargeNummer = Range("ChargenummerValue").Value

  ' Go to Annahme page
  Sheet8.Activate
  Sheet8.Range("ChargenummerCell").Value = chargeNummer
End Sub

Private Sub Prozess_Btn_Click()
  Dim chargeNummer
  chargeNummer = Range("ChargenummerValue").Value

  ' Go to prozess page
  Sheet1.Activate
  Sheet1.Range("ChargenummerCell").Value = chargeNummer
End Sub

Private Sub Alterung_Btn_Click()
  Dim chargeNummer
  chargeNummer = Range("ChargenummerValue").Value

  ' Go to alterung page
  Sheet5.Activate
  Sheet5.Range("ChargenummerCell").Value = chargeNummer
End Sub

Private Sub Nacharbeit_Btn_Click()
  Dim chargeNummer
  chargeNummer = RangE("ChargenummerValue").Value

  ' Go to nacharbeit page
  Sheet7.Activate
  Sheet7.Range("ChargenummerCell").Value = chargeNummer
End Sub

Private Sub Scrap_Btn_Click()
' Open the database file in hidden mode, search and modify the specified Charge.
' It will change the Chargephase on the Chargenummer only if it can find id

  ' Definition of variables
  Dim app As New excel.Application
  Dim wbook As excel.Workbook
  Dim sheet As Worksheet
  Dim table As ListObject
  Dim charges As Variant
  Dim i, j As Long
  Dim chargeRow As Integer
  Dim databasePath As String
  Dim chargeNummer As String
  databasePath = Sheet10.Range("ProjectPfad").Value & "Q_Silber_DATABASE.xlsx"

  If MsgBox(str_confirmScrapping, vbCritical + vbYesNo, str_notify_warningTitle) = vbYes Then

    ' Show saving message
    Application.StatusBar = str_statusBar_DBOpening

    ' Check if database is open
    For j = 1 To 30
      If IsDatabaseOpen(databasePath) Then
        Application.StatusBar = str_statusBar_DBBusy
        If j = 30 Then
          Notify str_notify_errorTitle, str_notify_scrappingDatabaseBusy, Sheet10.Range("ProjectPfad").Value, "Error", 20
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

    ' Searching the row to modify
    charges = table.DataBodyRange
    For i = LBound(charges) To UBound(charges)
      If CStr(charges(i, 1)) = CStr(Range("ChargenummerValue").Value) Then
        chargeRow = i
        Exit For
      End If
    Next i

    ' Write data in database
    With table
      .DataBodyRange(chargeRow, QSilberDB_Col.Phase).Value = ChargePhase.Scrapped
    End If

    ' Save and close the database
    wbook.SaveAs databasePath
    app.Quit
    Set app = Nothing

    ' Show confirmation message
    Application.StatusBar = StringFormat(str_statusBar_chargeScrapped, Range("ChargenummerValue").Value)
    Notify Range("ChargenummerValue").Value, str_notify_chargeScrappedText, Sheet10.Range("ProjectPfad").Value, "Info", 10

    ' Reset the form
    chargeNummer = Range("ChargenummerValue").Value
    FormReset

    ' Refresh all data connections & queries
    ThisWorkbook.RefreshAll
    DoEvents

    ' Reload data
    Range("ChargenummerValue").Value = chargeNummer
  End If
End Sub

Private Sub Delete_Btn_Click()
' Open the database file in hidden mode and delete the specified Charge.
' It will delete a row only if it can find it

  ' Definition of variables
  Dim app As New excel.Application
  Dim wbook As excel.Workbook
  Dim sheet As Worksheet
  Dim table As ListObject
  Dim charges As Variant
  Dim i, j As Long
  Dim chargeRow As Integer
  Dim databasePath As String
  databasePath = Sheet10.Range("ProjectPfad").Value & "Q_Silber_DATABASE.xlsx"
  Dim schichtcode As String
  schichtcode = CStr(Range("Schichtcode").Value)

  If MsgBox(str_confirmDeletion, vbCritical + vbYesNo, str_notify_attentionTitle) = vbYes Then

    ' Show saving message
    Application.StatusBar = str_statusBar_DBOpening

    ' Check if database is open
    For j = 1 To 30
      If IsDatabaseOpen(databasePath) Then
        Application.StatusBar = str_statusBar_DBBusy
        If j = 30 Then
          Notify str_notify_errorTitle, str_notify_deletingDatabaseBusy, Sheet10.Range("ProjectPfad").Value, "Error", 20
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

    ' Searching the row to delete
    charges = table.DataBodyRange
    For i = LBound(charges) To UBound(charges)
      If CStr(charges(i, 1)) = CStr(Range("ChargenummerValue").Value) Then
        chargeRow = i
        Exit For
      End If
    Next i

    ' Delete the row
    table.ListRows(chargeRow).Delete

    ' Save and close the database
    wbook.SaveAs databasePath
    app.Quit
    Set app = Nothing

    ' Check shiftcounter
    If Range("Schichtzähler").Value = 1 Then
      CheckShiftcounter schichtcode
    End If

    ' Show confirmation message
    Application.StatusBar = StringFormat(str_statusBar_chargeDeleted, Range("ChargenummerValue").Value)
    Notify Range("ChargenummerValue").Value, str_notify_chargeDeletedText, Sheet10.Range("ProjectPfad").Value, "Info", 10

    ' Reset the form
    FormReset

    ' Refresh all data connections & queries
    ThisWorkbook.RefreshAll
  End If
End Sub

Private Sub FormReset()
' Resets the form, deleting all the data
  Set sheet = Sheet6
  
  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword

  ' Reset form
  Range("ChargenummerCell").ClearContents
  Range("ChargenummerCell").Select
  ActiveWindow.Scrollrow = 1
  ActiveWindow.ScrollColumn = 1
  
  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
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

Function CheckShiftcounter(schichtcode As String)
' Control if the deleted row leave behind a wrong shift counter

  ' Definition of variables
  Dim app As New excel.Application
  Dim wbook As excel.Workbook
  Dim sheet As Worksheet
  Dim table As ListObject
  Dim charges As Variant
  Dim i, j As Long
  Dim chargeRow As Integer
  Dim databasePath As String
  databasePath = Sheet10.Range("ProjectPfad").Value & "Q_Silber_DATABASE.xlsx"

  ' Show message
  Application.StatusBar = "Schichten werden gerade überprüft ..."

  ' Check if database is open
  For j = 1 To 30
    If IsDatabaseOpen(databasePath) Then
      Application.StatusBar = str_statusBar_DBBusy
      If j = 30 Then
        Notify str_notify_errorTitle, str_notify_deletingDatabaseBusy, Sheet10.Range("ProjectPfad").Value, "Error", 20
        Exit Function
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

  ' Check Shiftcounter
  charges = table.DataBodyRange
  For i = LBound(charges) To UBound(charges)
    If CStr(charges(i, QSilberDB_Col.Schichtcode)) = schichtcode Then
      chargeRow = i
      Exit For
    End If
  Next i

  ' Update the Schichtzähler
  table.DataBodyRange(chargeRow, QSilberDB_Col.Schichtzähler).Value = 1

  ' Save and close the database
  wbook.SaveAs databasePath
  app.Quit
  Set app = Nothing
End Function