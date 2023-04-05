Option Explicit

' Global variables
Global Const GAdminPassword As String = "galvanik2023"

' Charge phases
Public Enum ChargePhase
  ' After the registration and creation of row
  Processing = 1

  ' Done testing, testing data added (layers), charge needs rework
  Nacharbeit

  ' Done testing, needs to be welded
  Alterung

  ' Charge is done and booked
  Done

  ' Scrapped, not booked
  Scrapped
End Enum

' Nacharbeit types
Public Enum NacharbeitTyp
  ' Too less silver layer, parts need to go back to the silver plant
  Dicke = 1
  
  ' There are spots on the parts, parts can be polished to remove the spots
  Flecken_A13

  ' There are spots on the parts, parts can be silvered one more time
  Flecken_EZ10

  ' There is too much silver on the layer, parts need to be stripped off the silver
  Strippen
End Enum

' Qualitätsaufzeichnung Database columns
Public Enum QSilberDB_Col
  ChrgNummer = 1
  Annahme_Datum
  Prozess_Datum
  Wochentag
  KW
  Monat
  Schicht
  Annahme_Mitarbeiter
  Annhame_Kommentar
  Materialnummer
  Gewicht_netto
  CuSchicht_soll
  AgSchicht_soll
  Auftragsnummer
  Füllmenge
  ChargeGewicht
  Stückzahl
  AgBedarf_soll
  CuBedarf_soll
  Anlage
  Trommel
  CuWert_soll
  CuWert_ist
  CuStrom_soll
  CuStrom_ist
  AgWert_soll
  AgWert_ist
  AgStrom_soll
  AgStrom_ist
  Prozess_Mitarbeiter
  Prozess_Kommentar
  Nacharbeit_Art
  Nacharbeit_Name
  Nacharbeit_Kommentar
  Nacharbeit_Mitarbeiter
  Nacharbeit_Anlage
  Nacharbeit_Trommel
  Nacharbeit_AgWert_soll
  Nacharbeit_AgWert_ist
  Nacharbeit_Kosten
  CuSchicht_ist
  AgSchicht_ist
  AgBedarf_ist
  AgSchicht_Nacharbeit
  AgBedarf_total
  AgBedarf_max
  Biegetest_iO
  AgEingespart
  Alterung_iO
  Alterung_Mitarbeiter
  Alterung_Datum
  Qualität_Kommentar
  Phase
  Phase_Name
  Schichtcode
  Schichtzähler
  Schicht_Datum
  Cu_UQG
  Cu_OQG
  Ag_UQG
  Ag_OQG
End Enum

' Teiledatenbank Database columns
Public Enum TeilDB_Col
  Materialnummer = 1
  Bezeichnung
  Grundmaterial
  Gewicht_netto
  Anlage_soll
  CuSchicht
  CuBedarf
  CuEZ
  CuStrom
  AgSchicht
  AgBedarf
  AgEZ_Straße
  AgStrom_Straße
  AgEZ_Glocke
  AgStrom_Glocke
  Löten
  XRay_Methode
  Kommentar
  Erstelldatum
  Ersteller
  Änderungsdatum
  Editor
End Enum

Public Function Notify(ByVal title As String, ByVal msg As String, ByVal iconPath As String, _
                    Optional ByVal notification_icon As String = "Info", _
                    Optional ByVal duration As Integer = 10)
' This public function sends notification using Windows 10 Notification API
' It uses a fast powershell command to do that using System.Windows.Forms
' Available parameters:
'    title (str): Notification title
'    msg (str): Notification message
'    notification_icon (str): Notification icon. Available options are: Info, Error and Warning
'    duration (int): Duration of notification in seconds, default is 10
  
  Dim notifyIcon As String
  Dim WsShell As Object: Set WsShell = CreateObject("WScript.Shell")
  Dim strCommand  As String

  ' Sanify notification_icon parameter
  If notification_icon <> "Info" And notification_icon <> "Error" And notification_icon <> "Warning" Then
    notification_icon = "Info"
  End If

  ' Build notification object
  notifyIcon = iconPath & "\Icons\notification.ico"
  strCommand = "powershell.exe -Command " & Chr(34) & "& { "
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

Public Function StringFormat(ByVal mask As String, ParamArray tokens()) As String
' This public function can be used to simplify the process of generating strings
' with dynamic content by allowing the use of placeholders in the string and
' replacing them with actual values during runtime.
  Dim i As Long
  
  For i = LBound(tokens) To UBound(tokens)
    mask = Replace(mask, "{" & i & "}", tokens(i))
  Next
  
  StringFormat = mask
End Function