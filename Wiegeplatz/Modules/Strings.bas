' --------------------------------------------
' MODULE FOR STRINGS LOCALIZATION
' Here are all the messages used throught the
' project.
' --------------------------------------------

' Common strings
Public Const str_materialNotFound = "Teilenummer nicht vorhanden!" & vbCrLf & "Jetzt hinzufügen?"
Public Const str_invalidBarcode = "Ungültiger barcode! Bitte ernuet scannen ..."

Public Const str_statusBar_DBOpening = "Datenbank wird geöffnet, bitte warten ..."
Public Const str_statusBar_DBBusy = "Datenbank beschäftigt, bitte warten ..."
Public Const str_statusBar_DBSaving = "Wird gespeichert ..."
Public Const str_statusBar_Printing = "Wird ausgedruckt ..."

Public Const str_notify_errorTitle = "FEHLER"
Public Const str_notify_attentionTitle = "ACHTUNG"

Public Const str_notify_savingDatabaseBusy = "Charge wurde NICHT gespeichert, weil die Datenbank ausgelastet ist! Bitte ernuet probieren"

Public Const str_saving_disclaimer = "Bitte speichern Sie diese Datei NICHT. Das Spechern der Datei wird nicht empfohlen!" & vbCrLf & "Mit dem Speichern fortfahren?"

' Sheet 1
Public Const str_statusBar_newChargeSaved = "Neue Charge für Material {0} erfolgreich gespeichert"
Public Const str_notify_newChargeSavedText = "Charge erfolgreich hinzugefügt"
Public Const str_kisteInDB = "Diese Kiste wurde bereits angemeldet!" & vbCrLf & "Wenn Sie fortfahren, wird eine weitere Charge (DOPPELT) erstellt." & vbCrLf & vbCrLf & "Wenn Sie Probleme mit dem Chargeschein hatten, sollten Sie das Chargeschein nur nachdrucken und KEINEN weiteren Charge erstellen." & vbCrLf & vbCrLf & "Wollen Sie wirklich weitermachen und eine doppelte Charge erstellen?"

' Sheet 4
Public Const str_notify_savingNewMaterialDatabaseBusy = "Material wurde NICHT gespeichert, weil die Datenbank ausgelastet ist! Bitte ernuet probieren"

Public Const str_statusBar_materialEdited = "Material-Nr. {0}} wurde aktualisiert"
Public Const str_notify_materialEditedText = "Teile erfolgreich aktualisiert"

Public Const str_statusBar_newMaterialAdded = "Material-Nr. {0} wurde erstellt"
Public Const str_notify_newMaterialAddedText = "Teile erfolgreich hinzugefügt"

Public Const str_confirmDeletion = "Möchten Sie dieses Material wirklich dauerhaft löschen?"
Public Const str_notify_deletingMaterialDatabaseBusy = "Material wurde NICHT gelöscht, weil die Datenbank ausgelastet ist! Bitte erneut probieren"

Public Const str_statusBar_materialDeleted = "Material-Nr. {0} wurde gelöscht"
Public Const str_notify_materialDeletedText = "Material erfolgreich gelöscht"