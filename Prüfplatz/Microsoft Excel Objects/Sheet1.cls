Option Explicit

Public sheet As Worksheet

Private Sub Worksheet_Activate()
  FormReset True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
' Is triggered everytime a cell content changes in the Worksheet

  ' Control if barcode cell was changed and act accordingly
  If Not Intersect(Target, Range("Prozess_Barcode")) Is Nothing Then
    If Not Range("Prozess_IsBarcodeEmpty").Value Then
      If Range("Prozess_IsBarcodeValid").Value Then
        If Range("Prozess_IsChargeInDB").Value Then
          If Range("Prozess_Phase").Value = 0 Then
            FormReset False
            CreateMode_Fill
          Else
            FormReset False
            EditMode_Fill
          End If
        Else
          MsgBox "Die Charge wurde nicht in der Datenbank gefunden!"
          FormReset True
        End If
      Else
        MsgBox "Ungültiger barcode!"
        FormReset True
      End If
    End If
  End If
End Sub

Private Sub FormReset(ByVal FullReset As Boolean)
' Resets the form, deleting all the custom data and moving the selection
' to the barcode Cell

  Set sheet = Sheet1

  ' Disable protection
  sheet.Unprotect Password:=GAdminPassword

  ' Reset form
  If FullReset Then
    Range("Prozess_Passivierung").ClearContents
    Range("Prozess_CuZahlerIst").ClearContents
    Range("Prozess_AgZahlerIst").ClearContents
    Range("Prozess_AgStromIst").ClearContents
    Range("Prozess_Barcode").ClearContents
  End If
  Range("Prozess_Trommel").ClearContents
  Range("Prozess_Kommentar").ClearContents
  Range("Prozess_Mitarbeiter").ClearContents

  Range("Prozess_Barcode").Select

  ' Enable protection
  sheet.Protect Password:=GAdminPassword
  sheet.EnableSelection = xlUnlockedCells
End Sub

Private Sub CreateMode_Fill()
' Compile the form in creation mode. Copies only MA number from
' the Qualitätsdatabase, because normally the "Prüfer" is the same
' of the process

  ' Copy MA Number from Q_DB
  If CSng(Range("I21").Value) <> 0 Then
    Range("Prozess_Mitarbeiter").Value = CSng(Range("I21").Value)
  End If
End Sub

Private Sub EditMode_Fill()
' Compile the form in edit mode. Copies the values from the Quality
' Database to be modified by the Worker and then saved.
' Warning: values from database are retrieved by formulas in the Excel
' sheet and not by the macro. They are referenced as normal range
' pointing to the cells.

  ' Copy "Trommel"
  If CStr(Range("I10").Value) <> "" Then
    Range("Prozess_Trommel").Value = CStr(Range("I10").Value)
  End If

  ' Copy "Passivierungsbad"
  If CStr(Range("I11").Value) <> "" Then
    Range("Prozess_Passivierung").Value = CStr(Range("I11").Value)
  End If

  ' Copy "Kupferdaten"
  If CSng(Range("I13").Value) > 0 Then
    Range("Prozess_CuZahlerIst").Value = CSng(Range("I13").Value)
  End If
  If CSng(Range("I15").Value) > 0 Then
    Range("Prozess_CuSchichtIst").Value = CSng(Range("I15").Value)
  End If

  ' Copy "Silberdaten"
  If CSng(Range("I17").Value) > 0 Then
    Range("Prozess_AgZahlerIst").Value = CSng(Range("I17").Value)
  End If
  If CSng(Range("I18").Value) > 0 Then
    Range("Prozess_AgStromIst").Value = CSng(Range("I18").Value)
  End If
  If CSng(Range("I19").Value) > 0 Then
    Range("Prozess_AgSchichtIst").Value = CSng(Range("I19").Value)
  End If

  ' Copy MA Number from Q_DB
  If CSng(Range("I21").Value) <> 0 Then
    Range("Prozess_Mitarbeiter").Value = CSng(Range("I21").Value)
  End If
End Sub
