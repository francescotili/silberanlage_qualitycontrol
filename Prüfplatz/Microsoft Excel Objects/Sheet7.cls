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
  SaveBtn.Enabled = False
  SaveBtn.BackColor = vbButtonFace
  If Range("IsFormValid").Value Then
    SaveBtn.Enabled = True
    SaveBtn.BackColor = RGB(0, 192, 0)
  End If

  ' Enable or disable specific form section based on the type of Nacharbeit
  If Not Intersect(Target, Range("NacharbeitString")) Is Nothing Then
    ' Disable protection
    sheet.Unprotect Password:=Sheet10.Range("DateiPassword").Value
    
    ' Change protection
    Range("Anlage").Locked = False
    Range("Trommel").Locked = False
    Range("Zählwert_ist").Locked = False
    Range("Schichtdicke").Locked = False
    If CStr(Range("NacharbeitArt").Value) <> "" Then
      Select Case Range("NacharbeitArt")
        Case NacharbeitTyp.Dicke
          Range("Anlage").ClearContents
        Case NacharbeitTyp.Flecken_A13
          Range("Anlage").Locked = True
          Range("Anlage").Value = "A13"
          Range("Trommel").Locked = True
          Range("Zählwert_ist").Locked = True
          Range("Schichtdicke").Locked = True
        Case NacharbeitTyp.Flecken_EZ10
          Range("Anlage").ClearContents
        Case NacharbeitTyp.Strippen
          Range("Anlage").Locked = True
          Range("Anlage").Value = "A13"
          Range("Trommel").Locked = True
          Range("Zählwert_ist").Locked = True
          Range("Schichtdicke").Locked = True
      End Select
    End If

    ' Enable protection
    sheet.Protect Password:=Sheet10.Range("DateiPassword").Value
    sheet.EnableSelection = xlUnlockedCells
  End If

End Sub

Private Sub EditMode_Fill()
' Compile the form in edit mode. Copies the values from the Quality
' Database to be modified by the Worker and then overwritten.
' Warning: values from database are retrieved by formulas in the Excel
' sheet and not by the macro. They are referenced as normal range
' pointing to the cells.

  ' Disable protection
  sheet.Unprotect Password:=Sheet10.Range("DateiPassword").Value

  ' Copy Nacharbeit comments
  If CStr(Range("KommentarDB").Value) <> "" Then
    Range("Kommentar").Value = Range("KommentarDB").Value
  End If

  ' Copy Anlage
  If CStr(Range("AnlageDB").Value) <> "" Then
    Range("Anlage").Value = Range("AnlageDB").Value
  End If

  ' Copy Trommel
  If CStr(Range("TrommelDB").Value) <> "" Then
    Range("Trommel").Value = Range("TrommelDB").Value
  End If

  ' Copy Silberwerte
  If CStr(Range("Zählerwert_istDB").Value) <> "" Then
    Range("Zählerwert_ist").Value = Range("Zählerwert_istDB").Value
  End If
  If CStr(Range("SchichtdickeDB").Value) <> "" Then
    Range("Schichtdicke").Value = Range("SchichtdickeDB").Value
  End If

  ' Copy Mitarbeiter
  If CStr(Range("MitarbeiterDB").Value) <> "" Then
    Range("Mitarbeiter").Value = Range("MitarbeiterDB").Value
  End If

  ' Copy Nacharbeit art ' LEAVE AT LAST POSITION BECAUSE IT CAUSES ERROR WITH SHEET PROTECTION
  If CStr(Range("NacharbeitArtDB").Value) <> "" Then
    Range("NacharbeitString").Value = Range("NacharbeitStringDB").Value
  End If

  ' Enable protection
  sheet.Protect Password:=Sheet10.Range("DateiPassword").Value
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub FormReset(ByVal FullReset As Boolean)
' Resets the form, deleting all the custom data and moving the selection
' to the barcode Cell

  Set sheet = Sheet7

  ' Disable protection
  sheet.Unprotect Password:=Sheet10.Range("DateiPassword").Value

  ' Reset form
  If FullReset Then
    Range("ChargenummerCell").ClearContents
  End If
  Range("NacharbeitString").ClearContents
  Range("Kommentar").ClearContents
  Range("Anlage").ClearContents
  Range("Trommel").ClearContents
  Range("Zählwert_ist").ClearContents
  Range("Schichtdicke").ClearContents
  Range("Mitarbeiter").ClearContents

  Range("ChargenummerCell").Select
  
  ' Enable protection
  sheet.Protect Password:=Sheet10.Range("DateiPassword").Value
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub CancelBtn_Click()
  FormReset True

  ' Go to main scanning page
  Sheet2.Activate
  Sheet2.Range("ChargenummerCell").Select
End Sub

Private Sub SaveBtn_Click()
' Save workflow triggered by Save button.

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
    .DataBodyRange(chargeRow, QSilberDB_Col.Nacharbeit_Art).Value = Range("NacharbeitArt").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.Nacharbeit_Kommentar).Value = Range("Kommentar").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.Nacharbeit_Mitarbeiter).Value = Range("Mitarbeiter").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.Nacharbeit_Anlage).Value = Range("Anlage").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.Nacharbeit_Trommel).Value = Range("Trommel").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.Nacharbeit_AgWert_soll).Value = Range("Zählerwert_soll").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.Nacharbeit_AgWert_ist).Value = Range("Zählwert_ist").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.AgSchicht_Nacharbeit).Value = Range("Schichtdicke").Value
    .DataBodyRange(chargeRow, QSilberDB_Col.Phase).Value = Range("NextChargePhase").Value
  End With

  ' Save and close the database
  wbook.SaveAs databasePath
  app.Quit
  Set app = Nothing

  ' Show confirmation message
  Application.StatusBar = "Charge für Auftrag " & Range("Auftragsnummer").Value & " erfolgreich ergänzt"
  Notify "Auftrag " & Range("Auftragsnummer").Value, "Charge erfolgreich ergänzt", "Info", 10

  ' Refresh all data connections & queries
  ThisWorkbook.RefreshAll

  ' Reset the form
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

Function ChargeWorkflow()
' The charge is in DB, decide what to do next

  If Range("IsEditMode").Value Then
    FormReset False
    EditMode_Fill
  Else
    FormReset False
  End If
End Function
