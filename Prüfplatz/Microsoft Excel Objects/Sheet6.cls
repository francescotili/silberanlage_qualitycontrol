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
            MsgBox "Die Charge wurde nicht in der Datenbank gefunden!"
            FormReset
          End If
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
  Sheet2.Range("ChargenummerCell").Select
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
  MsgBox "Verschrotten not yet implemented"
End Sub

Private Sub Delete_Btn_Click()
  MsgBox "Deletion not yet implemented"
End Sub

Private Sub FormReset()
' Resets the form, deleting all the data
  Set sheet = Sheet6
  
  ' Disable protection
  sheet.Unprotect Password:=Sheet10.Range("DateiPassword").Value

  ' Reset form
  Range("ChargenummerCell").ClearContents
  Range("ChargenummerCell").Select
  
  ' Enable protection
  sheet.Protect Password:=Sheet10.Range("DateiPassword").Value
  sheet.EnableSelection = xlUnlockedCells
End Sub
