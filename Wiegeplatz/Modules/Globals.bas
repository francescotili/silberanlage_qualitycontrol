Option Explicit

'Global variables
Global Const GAdminPassword As String = "galvanik2023"
Global Const GDatabasePath As String = "W:\X-Ray Qualitätsprüfung\Qualitätsaufzeichnung 2023_NEU\Q_Silber_DATABASE.xlsx"
Global Const GTeilDB_TableName As String = "Teiledatenbank"

' Qualitätsaufzeichnung Database columns
'   Chargenummer = 1
'   Datum
'   Wochentag
'   KW
'   Monat
'   Schicht
'   MitarbeiterAnnahme
'   KommentarAnnahme
'   Materialnummer
'   GewichtNetto
'   Kupferschicht_soll
'   Kupferschicht_min
'   Kupferschicht_max
'   Silberschicht_soll
'   Silberschicht_min
'   Silberschicht_max
'   Auftragsnummer
'   Füllmenge
'   Stückzahl
'   SilberBedarf_soll
'   KupferBedarf_soll
'   Anlage
'   Trommel
'   KupferLaufzeit_soll
'   KupferLaufzeit_ist
'   KupferStrom_soll
'   KupferStrom_ist
'   SilberLaufzeit_soll
'   SilberLaufzeit_ist
'   SilberStrom_soll
'   SilberStrom_ist
'   Passivierungsbad
'   KommentarProzess
'   Nacharbeit
'   Nacharbeit_Anlage
'   Nacharbeit_Trommel
'   Nacharbeit_Laufzeit_soll
'   Kupferschicht_ist
'   Silberschicht_ist
'   Silberschicht_nacharbeit
'   SilberBedarf_nacharbeit
'   Biegetest_iO
'   Biegetest_niO
'   SilberEingespart
'   Alterung_iO
'   Alterung_niO
'   Alterung_Mitarbeiter
'   Alterung_Datum
'   KommentarQualität
'   Entschied

' Teiledatenbank Database columns
'   Materialnummer = 1
'   Bezeichnung
'   Grundmaterial
'   GewichtNetto
'   Anlage_soll
'   Kupferschicht
'   KupferBedarf
'   Kupfer_EZ
'   KupferStrom
'   Silberschicht
'   SilberBedarf
'   Silber_EZ_Straße
'   SilberStrom_Straße
'   Silber_EZ_Glocke
'   SilberStrom_Glocke
'   Löten
'   XRayMethode
'   Kommentar

' EZ Datenbank Database columns
'   Materialnummer = 1
'   Grundmaterial
'   Silberschicht
'   Silber_EZ_Straße
'   Silber_EZ_Glocke
'   Kupferschicht
'   Kupfer_EZ
'   Kommentar

Public Function Notify(ByVal title As String, ByVal msg As String, _
                    Optional ByVal notification_icon As String = "Info", _
                    Optional ByVal duration As Integer = 10)
' This public function sends notification using Windows 10 Notification API
' Available parameters:
'    title (str): Notification title
'    msg (str): Notification message
'    notification_icon (str): Notification icon. Available options are: Info, Error and Warning
'    duration (int): Duration of notification in seconds, default is 10

  Const PSpath As String = "powershell.exe"
  Const notifyIcon As String = "W:\X-Ray Qualitätsprüfung\Qualitätsaufzeichnung 2023_NEU\Icons\success.ico"
  Dim WsShell As Object: Set WsShell = CreateObject("WScript.Shell")
  Dim strCommand  As String

  If notification_icon <> "Info" And notification_icon <> "Error" And notification_icon <> "Warning" Then
    notification_icon = "Info"
  End If

  ' Build notification object
  strCommand = """" & PSpath & """ -Command " & Chr(34) & "& { "
  strCommand = strCommand & "Add-Type -AssemblyName 'System.Windows.Forms'"
  strCommand = strCommand & "; $notification = New-Object System.Windows.Forms.NotifyIcon"
  strCommand = strCommand & "; $path = (Get-Process -id (get-process outlook).id[0]).Path"
  strCommand = strCommand & "; $notification.Icon = '" & notifyIcon & "'"
  strCommand = strCommand & "; $notification.BalloonTipIcon  = [System.Windows.Forms.ToolTipIcon]::" & notification_icon & ""
  strCommand = strCommand & "; $notification.BalloonTipText = '" & msg & "'"
  strCommand = strCommand & "; $notification.BalloonTipTitle = '" & title & "'"
  strCommand = strCommand & "; $notification.Visible = $true"
  strCommand = strCommand & "; $notification.ShowBalloonTip(" & duration & ")"
  strCommand = strCommand & " }" & Chr(34)

  ' Execute command, send notification
  WsShell.Run strCommand, 0, False
End Function

