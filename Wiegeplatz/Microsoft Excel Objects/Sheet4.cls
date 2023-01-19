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
Private silberBedarfCell, silberSchichtCell, anlageCell, barCodeCell, bezeichnungCell, kupferSchichtCell, alterungCell, grundMaterialCell, nettoGewichtCell As Range

Private Sub Worksheet_Activate()
  FormReset
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

