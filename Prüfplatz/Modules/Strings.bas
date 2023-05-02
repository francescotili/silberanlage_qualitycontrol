' --------------------------------------------
' MODULE FOR STRINGS LOCALIZATION
' Here are all the messages used throught the
' project.
' --------------------------------------------

' Common strings
Public Const str_chargeNotFound = "Die Charge wurde in der Datenbank NICHT gefunden!"
Public Const str_invalidBarcode = "Ungültiger barcode! Bitte ernuet scannen ..."
Public Const str_unhandledError = "Unerwartete Fehler!"

Public Const str_statusBar_DBOpening = "Datenbank wird geöffnet, bitte warten ..."
Public Const str_statusBar_DBBusy = "Datenbank beschäftigt, bitte warten ..."
Public Const str_statusBar_DBSaving = "Wird gespeichert ..."
Public Const str_statusBar_chargeSaved = "Charge für Auftrag {0} erfolgreich ergänzt"

Public Const str_notify_errorTitle = "FEHLER"
Public Const str_notify_warningTitle = "WARNUNG"
Public Const str_notify_attentionTitle = "ACHTUNG"

Public Const str_notify_savingDatabaseBusy = "Charge wurde NICHT gespeichert, weil die Datenbank ausgelastet ist! Bitte ernuet probieren"
Public Const str_notify_chargeSavedTitle = "Auftrag {0}"
Public Const str_notify_chargeSavedText = "Charge erfolgreich ergänzt"

Public Const str_saving_disclaimer = "Bitte speichern Sie diese Datei NICHT. Das Spechern der Datei wird nicht empfohlen!" & vbCrLf & "Mit dem Speichern fortfahren?"

' Sheet 5
Public Const str_chargeInProzess = "Diese Charge befindet sich noch in Bearbeitung, bitte zuerst die Prozess- und Qualitätsdaten hinzufügen!"
Public Const str_noLoetTest = "Dieses Teilnummer benötigt keine Löttest!"

' Sheet 6
Public Const str_confirmScrapping = "Möchten Sie diese Charge wirklich verschrotten?"
Public Const str_confirmDeletion = "Möchten Sie diese Charge wirklich dauerhaft löschen?"
Public Const str_notify_scrappingDatabaseBusy = "Charge wurde NICHT verschrottet, weil die Datenbank ausgelastet ist! Bitte ernuet probieren"
Public Const str_notify_deletingDatabaseBusy = "Charge wurde NICHT gelöscht, weil die Datenbank ausgelastet ist! Bitte ernuet probieren"

Public Const str_statusBar_chargeScrapped = "Charge-Nr. {0} wurde verschrottet"
Public Const str_notify_chargeScrappedText = "Charge erfolgreich verschrottet"

Public Const str_statusBar_chargeDeleted = "Charge-Nr. {0} wurde gelöscht"
Public Const str_notify_chargeDeletedText = "Charge erfolgreich gelöscht"