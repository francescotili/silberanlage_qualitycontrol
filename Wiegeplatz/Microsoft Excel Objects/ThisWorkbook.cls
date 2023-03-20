Private Sub Workbook_Open()
' Workaround to reset the form on the opening of the workbook

    ' Reset all the forms
    Application.ScreenUpdating = False
    Sheet2.Activate
    Sheet1.Activate
    Application.ScreenUpdating = True

    ' Refresh all data connections & queries
    ThisWorkbook.RefreshAll
    DoEvents
End Sub
