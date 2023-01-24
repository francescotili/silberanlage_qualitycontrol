Private Sub Workbook_Open()
' Workaround to reset the form on the opening of the workbook
    Application.ScreenUpdating = False
    Sheet2.Activate
    Sheet1.Activate
    Application.ScreenUpdating = True
End Sub
