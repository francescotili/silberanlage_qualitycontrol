Option Explicit

Public sheet As Worksheet
Private barCodeCell, barCodeValueCell, chargeWeightCell, anlageCell, isTeileInDBCell, saveBtnCell, resetBtnCell, isFormValidCell, isBarcodeValidCell, isBarcodeEmptyCell As Range
Private materialNummerInDB As Boolean

Private Sub Worksheet_Activate()
  FormReset
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
' Is triggered everytime a cell content changes in the Worksheet

  ' Define cell ranges
  Set barCodeCell = Range("NeueCharge_Barcode")
  Set isTeileInDBCell = Range("NeueCharge_IsTeileInDatabase")
  Set isFormValidCell = Range("NeueCharge_IsFormValid")
  Set isBarcodeValidCell = Range("NeueCharge_IsBarcodeValid")
  Set isBarcodeEmptyCell = Range("NeueCharge_IsBarcodeEmpty")
  
  ' Control if barcode cell was changed and act accordingly
  If Not Intersect(Target, barCodeCell) Is Nothing Then
    If Not isBarcodeEmptyCell.Value Then
      If isBarcodeValidCell.Value Then
        If Not isTeileInDBCell Then
          If MsgBox("Teilenummer nicht vorhanden!" & vbNewLine & "Jetzt hinzufügen?", vbInformation + vbYesNo + vbDefaultButton1, "ACHTUNG") = vbYes Then
            Dim barcode
            barcode = barCodeCell.Value
            Sheet4.Activate
            Sheet4.Range("NeueTeile_Barcode").Value = barcode
            Exit Sub
          End If
          FormReset
        End If
      Else
        MsgBox "Ungültiger barcode!"
        FormReset
      End If
    End If
  End If
  
  ' Enable or disable the save button
  NeueCharge_SpeichernBtn.Enabled = False
  NeueCharge_SpeichernBtn.BackColor = vbButtonFace
  If isFormValidCell.Value Then
    NeueCharge_SpeichernBtn.Enabled = True
    NeueCharge_SpeichernBtn.BackColor = RGB(0, 192, 0)
  End If
End Sub

Private Sub FormReset()
' Resets the form, deleting all the custom data and moving the selection
' to the barcode cell

  Set sheet = Sheet1
  
  ' Define cell ranges
  Set barCodeCell = Range("NeueCharge_Barcode")
  Set chargeWeightCell = Range("NeueCharge_Gewicht")
  Set anlageCell = Range("NeueCharge_Anlage")
  
  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword
  
  ' Reset form
  barCodeCell.ClearContents
  chargeWeightCell.ClearContents
  anlageCell.Formula = "=I16"
  barCodeCell.Select
  
  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub NeueCharge_SpeichernBtn_Click()
  MsgBox "It works!"
  FormReset
End Sub

Private Sub NeueCharge_AbbrechenBtn_Click()
  FormReset
End Sub
