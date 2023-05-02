Private Sub ResetDB_Btn_Click()
' Prepare the database for a new year of production data
' It basically delete the production data table (Qualitätsdatabase)
' It leaves the material data table and EZ table untouched
'
' In this way, the workers still have all the material data, but
' the production data will be emptied.
  
  Dim QTable As ListObject

  ' Ask for confirmation
  If MsgBox("Diese Funktion bereitete die Datenbank auf den Start ins neue Jahr vor." & vbCrLf & "Die Qualitätsdatenbank (Produktionsdaten) wird gelöscht! Der Materialdatenbank wird nicht berührt." & vbCrLf & vbCrLf & "Sind Sie wirklich sicher, dass Sie mit dem Zurückstezen fortfahren möchten?", vbExclamation + vbYesNo + vbDefaultButton2, "ACHTUNG") = vbYes Then

    ' Delete the data
    Application.ScreenUpdating = False
    Set QTable = Sheet1.ListObjects.Item("Qualitätsdatabase")
    With QTable
      If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Delete
      End If
    End With

    ' Visualize success message
    Application.ScreenUpdating = True
    Application.StatusBar = "Database erfolgreich zurückgesetzt !!"
    MsgBox("Die Datenbank wurde erfolgreich zurückgesetzt!")
  End If
End Sub