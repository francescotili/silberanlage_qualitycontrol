Option Explicit

Public sheet As Worksheet

Private Sub Worksheet_Activate()
  FormReset
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
' Is triggered everytime a cell content changes in the Worksheet

  ' Control if barcode cell was changed and act accordingly
  If Not Intersect(Target, Range("Info_Chargenummer")) Is Nothing Then
    If Not Range("Info_IsBarcodeEmpty").Value Then
      If Range("Info_IsBarcodeValid").Value Then
        If Range("Info_IsChargeInDB").Value Then
          'Form is automatic filled
        Else
          MsgBox "Die Charge wurde nicht in der Datenbank gefunden!"
          FormReset
        End If
      Else
        MsgBox "Ungültiger barcode!"
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
  If Range("Info_IsChargeInDB").Value Then
    Annahme_Btn.Enabled = True
    Prozess_Btn.Enabled = True
    If Range("L23") Then
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
  Sheet2.Range("Prufplatz_Chargenummer").Select
End Sub

Private Sub Annahme_Btn_Click()
  MsgBox "Annahme not yet implemented"
End Sub

Private Sub Prozess_Btn_Click()
  Dim chargeNummer
  chargeNummer = Range("Info_Chargenummer").Value

  ' Go to prozess page
  Sheet1.Activate
  Sheet1.Range("Prozess_Chargenummer").Value = chargeNummer
End Sub

Private Sub Alterung_Btn_Click()
  Dim chargeNummer
  chargeNummer = Range("Info_Chargenummer").Value

  ' Go to alterung page
  Sheet5.Activate
  Sheet5.Range("Alterung_Chargenummer").Value = chargeNummer
End Sub

Private Sub Nacharbeit_Btn_Click()
  MsgBox "Nacharbeit not yet implemented"
End Sub

Private Sub Scrap_Btn_Click()
  MsgBox "Verschrotten not yet implemented"
End Sub

Private Sub Delete_Btn_Click()
  MsgBox "Deletion not yet implemented"
End Sub

Private Sub FormReset()
' Resets the form, deleting all the data
  Set sheet = Sheet6
  
  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword

  ' Reset form
  Range("Info_Chargenummer").ClearContents
  Range("Info_Chargenummer").Select
  
  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub
