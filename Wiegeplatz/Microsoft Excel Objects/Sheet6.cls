Option Explicit

Private Sub Worksheet_Activate()
  FormReset
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
' Is triggered everytime a cell content changes in the worksheet
  
  ' Enable or disable the save button
  PrintBtn.Enabled = False
  PrintBtn.BackColor = vbButtonFace
  If Range("IsFormValid").Value Then
    PrintBtn.Enabled = True
    PrintBtn.BackColor = RGB(0, 192, 0)
  End If
End Sub

Private Sub FormReset()
' Resets the form, deleting all the custom data and moving the selection
' to the barcode cell

  ' Reset form
  Range("Auftrag").ClearContents
  Range("Kistenr").ClearContents
  
  ' Reset the view
  Range("Auftrag").Select
  ActiveWindow.ScrollRow = 1
  ActiveWindow.ScrollColumn = 1
End Sub

Private Sub PrintBtn_Click()
  Sheet1.Activate
  Sheet1.RePrint CStr(Range("Barcode").Value), CSng(Range("ChargeGewicht").Value), CStr(Range("Chargenummer").Value)
End Sub

Private Sub CancelBtn_Click()
  FormReset
End Sub
