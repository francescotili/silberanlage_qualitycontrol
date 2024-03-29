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
        If Range("IsChargeInDB").Value Then
          ChargeWorkflow
        Else
          ThisWorkbook.RefreshAll
          DoEvents ' Waits for the refresh to finish
          If Range("IsChargeInDB").Value Then
            ChargeWorkflow
          Else
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
End Sub

Private Sub FormReset()
' Resets the form, deleting all the custom data and moving selection
' to the barcode cell
  Set sheet = Sheet2
  
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

Function ChargeWorkflow()
  Dim chargeNummer
  chargeNummer = Range("ChargenummerValue").Value

  Select Case Range("CurrentChargePhase")
    Case ChargePhase.Processing ' Process and quality data are missing
      Sheet1.Activate
      Sheet1.Range("ChargenummerCell").Value = chargeNummer
      Exit Function
    Case ChargePhase.Nacharbeit ' Nacharbeit done, choose based on nacharbeit type
      Sheet7.Activate
      Sheet7.Range("ChargenummerCell").Value = chargeNummer
    Case ChargePhase.Alterung ' Alterung done
      Sheet5.Activate
      Sheet5.Range("ChargenummerCell").Value = chargeNummer
      Exit Function
    Case ChargePhase.Done To ChargePhase.Scrapped ' Charge is done or scrapped
      Sheet6.Activate
      Sheet6.Range("ChargenummerCell").Value = chargeNummer
      Exit Function
    Case Else
      MsgBox str_unhandledError
      FormReset
  End Select
End Function
