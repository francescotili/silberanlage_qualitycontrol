Private Sub Workbook_Open()
' Workaround to reset the form on the opening of the workbook

    ' Reset all the forms
    Application.ScreenUpdating = False
    Sheet1.Activate
    Sheet5.Activate
    Sheet6.Activate
    Sheet7.Activate
    Sheet8.Activate
    Sheet2.Activate ' START
    Application.ScreenUpdating = True

    ' Refresh all data connections & queries
    ThisWorkbook.RefreshAll
    DoEvents
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  If MsgBox(str_saving_disclaimer, vbExclamation + vbYesNo + vbDefaultButton2, str_notify_warningTitle) = vbNo Then
    Cancel = True
  End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  ActiveWorkbook.Saved = True
End Sub