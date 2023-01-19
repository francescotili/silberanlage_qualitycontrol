VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
' Workaround to reset the form on the opening of the workbook
    Application.ScreenUpdating = False
    Sheet2.Activate
    Sheet1.Activate
    Application.ScreenUpdating = True
End Sub
