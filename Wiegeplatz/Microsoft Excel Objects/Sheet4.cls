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
  If CLng(Range("I13").Value) <> 0 Then
    kupferSchichtCell.Value = Range("I13").Value
  End If
  
  ' Copy "Silberschicht"
  Set silberSchichtCell = Range("NeueTeile_AgSchicht")
  If CLng(Range("I14").Value) <> 0 Then
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
  If CLng(Range("I21").Value) > 0 Then
    CuEZCell.Value = Range("I21").Value
  End If
  
  ' Copy Ag Data from DB
  Set AgStrasseEZCell = Range("NeueTeile_AgEZStraße")
  If CLng(Range("I24").Value) > 0 Then
    AgStrasseEZCell.Value = Range("I24").Value
  End If
  Set AgStrasseStromCell = Range("NeueTeile_AgStromStraße")
  If CLng(Range("I25").Value) > 0 Then
    AgStrasseStromCell.Value = Range("I25").Value
  End If
  
  Set AgGlockeEZCell = Range("NeueTeile_AgEZGlocke")
  If CLng(Range("I26").Value) > 0 Then
    AgGlockeEZCell.Value = Range("I26").Value
  End If
  Set AgGlockeStromCell = Range("NeueTeile_AgStromGlocke")
  If CLng(Range("I27").Value) > 0 Then
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
  If CLng(Range("K21").Value) > 0 Then
    CuEZCell.Value = Range("K21").Value
  End If
  
  ' Copy Ag Data from DB
  Set AgStrasseEZCell = Range("NeueTeile_AgEZStraße")
  If CLng(Range("K24").Value) > 0 Then
    AgStrasseEZCell.Value = Range("K24").Value
  End If
  
  Set AgGlockeEZCell = Range("NeueTeile_AgEZGlocke")
  If CLng(Range("K26").Value) > 0 Then
    AgGlockeEZCell.Value = Range("K26").Value
  End If
End Sub

Private Sub NeueTeile_SpeichernBtn_Click()
  MsgBox "It works!"
  FormReset True
End Sub

Private Sub NeueTeile_LoeschenBtn_Click()
  If MsgBox("Möchten Sie dieses Material wirklich dauerhaft löschen?", vbCritical + vbYesNo, "Achtung") = vbYes Then
    MsgBox "Gelöscht"
    FormReset True
  End If
End Sub

Private Sub NeueTeile_ResetBtn_Click()
  FormReset True
End Sub
