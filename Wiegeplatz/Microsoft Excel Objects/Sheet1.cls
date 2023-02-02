Option Explicit

Public sheet As Worksheet
Private barCodeCell, barCodeValueCell, chargeWeightCell, anlageCell, isTeileInDBCell, saveBtnCell, resetBtnCell, isFormValidCell, isBarcodeValidCell, isBarcodeEmptyCell As Range
Private materialNummerInDB As Boolean

Private Sub Worksheet_Activate()
  FormReset True
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
        If isTeileInDBCell Then
          FormReset False
        Else
          If MsgBox("Teilenummer nicht vorhanden!" & vbNewLine & "Jetzt hinzufügen?", vbInformation + vbYesNo + vbDefaultButton1, "ACHTUNG") = vbYes Then
            Dim barcode
            barcode = barCodeCell.Value
            Sheet4.Activate
            Sheet4.Range("NeueTeile_Barcode").Value = barcode
            Exit Sub
          End If
          FormReset True
        End If
      Else
        MsgBox "Ungültiger barcode!"
        FormReset True
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

Private Sub FormReset(ByVal BarcodeReset As Boolean)
' Resets the form, deleting all the custom data and moving the selection
' to the barcode cell

  Set sheet = Sheet1
  
  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword
  
  ' Reset form
  If BarcodeReset Then
    Set barCodeCell = Range("NeueCharge_Barcode")
    barCodeCell.ClearContents
  End If
  
  Set chargeWeightCell = Range("NeueCharge_Gewicht")
  chargeWeightCell.ClearContents
  
  Set anlageCell = Range("NeueCharge_Anlage")
  anlageCell.ClearContents
  
  Set kupferSchichtCell = Range("NeueCharge_CuSchicht")
  kupferSchichtCell.ClearContents
  
  Set silberSchichtCell = Range("NeueCharge_AgSchicht")
  silberSchichtCell.ClearContents
  
  barCodeCell.Select
  
  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub NeueCharge_SpeichernBtn_Click()
  MsgBox "It works!"
  FormReset True
End Sub

Private Sub NeueCharge_AbbrechenBtn_Click()
  FormReset True
End Sub
