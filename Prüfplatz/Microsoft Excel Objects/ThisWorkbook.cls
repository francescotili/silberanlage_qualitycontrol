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

'Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'  MsgBox "Das Speichern dieser Datei ist deaktiviert, bitte schließen es ohne zu speichern!"
'  Cancel = True
'End Sub
'
'Private Sub Workbook_BeforeClose(Cancel As Boolean)
'  ActiveWorkbook.Saved = True
'End Sub