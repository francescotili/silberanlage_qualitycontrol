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