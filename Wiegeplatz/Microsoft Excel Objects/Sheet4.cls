VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public sheet As Worksheet
Private silberBedarfCell, silberSchichtCell, anlageCell, barCodeCell, bezeichnungCell As Range
Private kupferSchichtCell, alterungCell, grundMaterialCell, nettoGewichtCell As Range
Private barCodeValueCell, isTeileInDBCell, isFormValidCell, isBarcodeValidCell As Range
Private CuEZCell, AgEZCell, AgStrom, isBarcodeEmptyCell As Range
Private AgStrasseEZCell, AgGlockeEZCell, AgStrasseStromCell, AgGlockeStromCell As Range

Private Sub Worksheet_Activate()
  FormReset
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
          ' Edit Mode
          MsgBox "Edit mode"
        Else
          ' Creation Mode
          MsgBox "Creation mode"
        End If
      Else
        MsgBox "Ungültiger barcode!"
        FormReset
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

Private Sub FormReset()
' Resets the form, deleting all the custom data and moving the selection
' to the barcode cell
  
  Set sheet = Sheet4
  
  ' Define cell ranges
  Set silberBedarfCell = Range("NeueTeile_AgBedarf")
  Set silberSchichtCell = Range("NeueTeile_AgSchicht")
  Set anlageCell = Range("NeueTeile_Anlage")
  Set barCodeCell = Range("NeueTeile_Barcode")
  Set bezeichnungCell = Range("NeueTeile_Bezeichnung")
  Set kupferSchichtCell = Range("NeueTeile_CuSchicht")
  Set alterungCell = Range("NeueTeile_Loettest")
  Set grundMaterialCell = Range("NeueTeile_Material")
  Set nettoGewichtCell = Range("NeueTeile_NettoGewicht")
  
  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword
  
  ' Reset form
  silberBedarfCell.ClearContents
  silberSchichtCell.ClearContents
  anlageCell.ClearContents
  barCodeCell.ClearContents
  bezeichnungCell.ClearContents
  kupferSchichtCell.ClearContents
  alterungCell.ClearContents
  grundMaterialCell.ClearContents
  nettoGewichtCell.ClearContents
  
  barCodeCell.Select
  
  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub NeueTeile_SpeichernBtn_Click()
  MsgBox "It works!"
  FormReset
End Sub

Private Sub NeueTeile_LoeschenBtn_Click()
  If MsgBox("Möchten Sie dieses Material wirklich dauerhaft löschen?", vbCritical + vbYesNo, "Achtung") = vbYes Then
    MsgBox "Gelöscht"
  End If
End Sub