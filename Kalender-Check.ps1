<#
.SYNOPSIS
    Kalender-Check: shows for Exchange Server (on-prem) who may access which calendar, and reports some
    anomalies (default permission too broad, anonymous access, write permissions). The result is an HTML report.

    Deutsch:
    Kalender-Check: zeigt für Exchange Server (on-prem), wer auf welchen Kalender zugreifen darf, und meldet einige
    Auffälligkeiten (Standard-Recht zu weit, Anonym-Zugriff, Schreibrechte). Ergebnis ist ein HTML-Bericht.

.DESCRIPTION
    The script works exclusively READ-ONLY. It uses only Exchange Get commands (Get-Mailbox,
    Get-MailboxFolderStatistics, Get-MailboxFolderPermission), changes no permissions and transfers nothing.

    The main calendar of every mailbox is read. The folder is found by folder type, so the script works in any mailbox language; it reads sequentially and is correspondingly slow with very many mailboxes
    (roughly one second per mailbox). It is easiest to start directly on an Exchange server: The script loads the
    Exchange Management Shell itself.

    Prerequisites: Windows PowerShell 5.1, Exchange Server 2016/2019/SE, for reading, the role
    "View-Only Recipients" is sufficient (e.g. via "View-Only Organization Management").

    Deutsch:
    Das Skript arbeitet ausschließlich LESEND. Es verwendet nur Get-Befehle von Exchange (Get-Mailbox,
    Get-MailboxFolderStatistics, Get-MailboxFolderPermission), ändert keine Berechtigung und sendet nichts an Dritte oder ins Internet.

    Gelesen wird der Hauptkalender jedes Postfachs. Der Ordner wird über den Ordnertyp gesucht, deshalb funktioniert das
    Skript in jeder Postfachsprache; es liest nacheinander und ist bei sehr vielen Postfächern entsprechend langsam
    (grob eine Sekunde je Postfach). Am einfachsten direkt auf einem Exchange-Server starten: Das Skript lädt die
    Exchange-Verwaltungsshell selbst.

    Voraussetzungen: Windows PowerShell 5.1, Exchange Server 2016/2019/SE, zum Lesen genügt die Rolle
    "View-Only Recipients" (z. B. über "View-Only Organization Management").

.PARAMETER Ausgabe
    Path of the HTML file. Default: Kalender-Check_Exchange_<Datum>.html in the current folder.

    Deutsch:
    Pfad der HTML-Datei. Standard: Kalender-Check_Exchange_<Datum>.html im aktuellen Ordner.

.PARAMETER Postfach
    Only check these mailboxes (addresses). Default: all user, shared, room, and equipment mailboxes.

    Deutsch:
    Nur diese Postfächer prüfen (Adressen). Standard: alle Benutzer-, Gemeinsamen, Raum- und Gerätepostfächer.

.PARAMETER ExchangeServer
    Name or FQDN of an Exchange server for a remote session (Kerberos), if no Exchange management tools are installed on the computer.

    Deutsch:
    Name oder FQDN eines Exchange-Servers für eine Remote-Sitzung (Kerberos), wenn auf dem Rechner keine Exchange-
    Verwaltungstools installiert sind.

.PARAMETER Csv
    Additionally writes a CSV file for each check point with hits and the list of all calendar permissions alongside the report.

    Deutsch:
    Schreibt zusätzlich je Prüfpunkt mit Treffern eine CSV-Datei und die Liste aller Kalenderrechte neben den Bericht.

.PARAMETER NichtOeffnen
    Does not automatically open the finished report in the browser.

    Deutsch:
    Öffnet den fertigen Bericht nicht automatisch im Browser.

.PARAMETER PassThru
    Additionally returns the findings as objects (for custom evaluations).

    Deutsch:
    Gibt die Befunde zusätzlich als Objekte zurück (für eigene Auswertungen).

.PARAMETER Language
    Language of the report and the console output: 'de' (German) or 'en' (English). Default: German if the Windows display language
    is German, otherwise English.

    Deutsch:
    Sprache von Bericht und Konsolenausgabe: 'de' (Deutsch) oder 'en' (Englisch). Standard: Deutsch, wenn die Anzeigesprache von
    Windows Deutsch ist, sonst Englisch.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Kalender-Check.ps1

.EXAMPLE
    .\Kalender-Check.ps1 -ExchangeServer ex01.firma.local -Csv -Ausgabe C:\Berichte\kalender.html

.NOTES
    IT Admin Tools by MM - Kalender-Check Version 1.0 (11.2026), Lizenz: MIT (siehe LICENSE)
    Kein Ersatz für ein Berechtigungskonzept. Nutzung auf eigene Verantwortung.
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('de', 'en')][string]$Language = $(if ((Get-UICulture).TwoLetterISOLanguageName -eq 'de') { 'de' } else { 'en' }),
    [Alias('Output')][string]$Ausgabe,
    [string[]]$Postfach,
    [string]$ExchangeServer,
    [switch]$Csv,
    [Alias('NoOpen')][switch]$NichtOeffnen,
    [switch]$PassThru
)
# Sprache des Berichts: de oder en (Parameter -Language). T 'deutsch' 'english' liefert den Text der gewählten Sprache.
$En = ($Language -eq 'en')
function T([string]$de, [string]$en) { if ($script:En) { $en } else { $de } }

$ErrorActionPreference = 'Stop'
$Version = '1.1'

# Kommalisten: Beim Start über "powershell -File" kommt -Postfach a,b als ein Text an. Deshalb an Kommas trennen.
$nurAuswahl = $PSBoundParameters.ContainsKey('Postfach')
if ($Postfach) { $Postfach = @($Postfach | ForEach-Object { $_ -split '\s*,\s*' } | Where-Object { $_ }) }
if ($nurAuswahl -and -not $Postfach) { throw (T 'Bei -Postfach wurde kein Postfach angegeben. Ohne den Parameter werden alle Postfächer geprüft.' 'No mailbox was specified for -Postfach. Without the parameter, all mailboxes are checked.') }

#@SPRACHE@
# ------------------------------------------------------------------ Hilfsfunktionen
function Datum($d, $leer = $null) { if ($null -eq $leer) { $leer = T 'nie' 'never' }; if ($d) { ([datetime]$d).ToString((T 'dd.MM.yyyy' 'yyyy-MM-dd')) } else { $leer } }
function JaNein($b) { if ($b) { T 'Ja' 'Yes' } else { T 'Nein' 'No' } }
function Ort($dn) { ($dn -replace '^(CN|OU)=.+?(?<!\\),', '') -replace ',DC=.*$', '' }
function Html($t) { [System.Net.WebUtility]::HtmlEncode([string]$t) }
function Kurz($t, $n = 220) { $t = ([string]$t -replace '\s+', ' ').Trim(); if ($t.Length -gt $n) { $t.Substring(0, $n) + ' ...' } else { $t } }

# CSV-Dateien öffnen sich in Excel: Werte, die mit = + - @ beginnen, würden dort als Formel ausgeführt (Namen von Dateien, Ordnern, Konten oder Computern
# stammen von Benutzern). Diese Funktion ersetzt für dieses Skript den Befehl Export-Csv und setzt vor solche Werte ein Hochkomma.
function Export-Csv {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline = $true)]$InputObject, [Parameter(Mandatory)][string]$Path, [switch]$NoTypeInformation, [string]$Delimiter = ';', [string]$Encoding = 'UTF8')
    begin { $alle = New-Object System.Collections.Generic.List[object] }
    process { if ($null -ne $InputObject) { $alle.Add($InputObject) } }
    end {
        $sicher = foreach ($o in $alle) {
            $n = [ordered]@{}
            foreach ($p in $o.PSObject.Properties) { $v = $p.Value; if (($v -is [string]) -and ($v -match '^[=+@\t\r]|^-.')) { $v = "'" + $v }; $n[$p.Name] = $v }
            [pscustomobject]$n
        }
        if ($alle.Count) { $sicher | Microsoft.PowerShell.Utility\Export-Csv -Path $Path -NoTypeInformation:$NoTypeInformation -Delimiter $Delimiter -Encoding $Encoding }   # leere Liste: keine Datei (sie hätte keine Kopfzeile)
    }
}

# Tabellen: ConvertTo-Html und Export-Csv nehmen die Spalten des ersten Objekts. Haben die Zeilen einer Liste verschiedene
# Eigenschaften (z. B. Fehlzeile neben normalen Treffern), gingen Spalten verloren: alle Namen der Reihe nach vereinigen.
function Einheitlich($liste) {
    $liste = @($liste | Where-Object { $null -ne $_ })
    if ($liste.Count -ge 2) {
        $namen = New-Object System.Collections.Generic.List[string]
        foreach ($o in $liste) { foreach ($p in $o.PSObject.Properties.Name) { if (-not $namen.Contains($p)) { $namen.Add($p) } } }
        if ($namen.Count -ne @($liste[0].PSObject.Properties.Name).Count) { $liste = @($liste | Select-Object -Property $namen) }
    }
    SpaltenAnzeige $liste
}
# Englische Berichte: Spaltenüberschriften aus der Übersetzungstabelle (intern bleiben die deutschen Namen, der Code liest sie weiter)
$SpaltenEn = @{
    'Abfrage' = 'Query'
    'Administrator' = 'Administrator'
    'Adresse' = 'Address'
    'Aktiv' = 'Active'
    'Aktiv verknüpft mit' = 'Active linked to'
    'Aktive Verknüpfungen' = 'Active links'
    'Alter in Tagen' = 'Age in days'
    'Angelegt' = 'Created'
    'Anmeldung' = 'Last logon'
    'Anmeldungen' = 'Logons'
    'Anteil (%)' = 'Share (%)'
    'Anzahl' = 'Count'
    'Art' = 'Account type'
    'Auf Domänencontroller' = 'On domain controller'
    'Aufräumen' = 'Cleanup'
    'Auslastung' = 'Utilization'
    'Auslösender Rechner' = 'Triggering computer'
    'Auslösender Rechner (Anzahl)' = 'Triggering computers (count)'
    'Aussteller' = 'Issuer'
    'Beispiele' = 'Examples'
    'Belegt' = 'Used'
    'Belegt (%)' = 'Used (%)'
    'Benutzer und Computer (auch verschachtelt)' = 'Users and computers (also nested)'
    'Benutzereinstellungen' = 'User settings'
    'Benutzername' = 'Username'
    'Berechtigter' = 'Principal'
    'Bereich' = 'Scope'
    'Bereiche' = 'Scopes'
    'Beschreibung' = 'Description'
    'Besitzer' = 'Owner'
    'Betriebssystem' = 'OS'
    'Betroffene Konten' = 'Affected accounts'
    'Bewertung' = 'Rating'
    'Bis' = 'To'
    'BitLocker' = 'BitLocker'
    'Computer' = 'Computer'
    'Computereinstellungen' = 'Computer settings'
    'Darf anwenden' = 'Can apply'
    'Darf delegieren an' = 'Can delegate to'
    'Darf sich hier als Benutzer ausgeben' = 'Can log on as user'
    'Datei' = 'File'
    'Dateien' = 'Files'
    'Dateisystem' = 'File system'
    'Dateityp' = 'File type'
    'Datenbank' = 'Database'
    'Detail' = 'Detail'
    'Details' = 'Details'
    'Direkte Mitglieder' = 'Direct members'
    'Domäne' = 'Domain'
    'Domänencontroller' = 'Domain controller'
    'Domänennamenmaster' = 'Domain naming master'
    'Dynamische Updates' = 'Dynamic updates'
    'E-Mail' = 'Email'
    'Eigenschaft' = 'Property'
    'Einstellung' = 'Setting'
    'Empfohlen' = 'Recommended'
    'Enthaltene Gruppen' = 'Included groups'
    'Ereignisse' = 'Events'
    'Erreichbar' = 'Reachable'
    'Erste' = 'First'
    'Erster' = 'First'
    'Erzwungen' = 'Forced'
    'Failover' = 'Failover'
    'Falsches Kennwort' = 'Wrong password'
    'Fehlanmeldungen am DC' = 'Failed logons at DC'
    'Fehlversuche' = 'Failed attempts'
    'Feststellung' = 'Finding'
    'Fingerabdruck' = 'Fingerprint'
    'Firewall' = 'Firewall'
    'Frei' = 'Free'
    'Freigabe' = 'Share'
    'Freigaberecht' = 'Share permission'
    'Funktion' = 'Function'
    'Gelesen von' = 'Read by'
    'Geprüft' = 'Checked'
    'Gesamtgröße' = 'Total size'
    'Geschützte Benutzer' = 'Protected users'
    'Gesperrt seit' = 'Locked since'
    'Geändert' = 'Changed'
    'Gilt' = 'Applies to'
    'Gilt für' = 'Applies to'
    'Globaler Katalog' = 'Global catalog'
    'Grenze' = 'Limit'
    'Groesse' = 'Size'
    'Grund' = 'Reason'
    'Gruppe' = 'Group'
    'Gruppen' = 'Groups'
    'Gruppen (auch verschachtelt)' = 'Groups (including nested)'
    'Gruppenrichtlinie' = 'Group policy'
    'Größe' = 'Size'
    'Größe (GB)' = 'Size (GB)'
    'Gültig ab' = 'Valid from'
    'Gültig bis' = 'Valid to'
    'Herkunft' = 'Origin'
    'Hersteller' = 'Vendor'
    'Hinweis' = 'Note'
    'Häufigster Grund' = 'Most common reason'
    'IP-Adresse' = 'IP address'
    'Im AD gespeichert' = 'Stored in AD'
    'In der Domäne vorhanden' = 'Present in domain'
    'Infrastrukturmaster' = 'Infrastructure master'
    'Inhalt' = 'Content'
    'Ist' = 'Is'
    'Kennung' = 'ID'
    'Kennwort gesetzt' = 'Password set'
    'Kennwort läuft nie ab' = 'Password never expires'
    'Kerberos' = 'Kerberos'
    'Konten' = 'Accounts'
    'Konten (auch verschachtelt)' = 'Accounts (also nested)'
    'Konto' = 'Account'
    'Konto für DNS-Einträge' = 'Account for DNS entries'
    'Kopie bleibt im Postfach' = 'Copy remains in mailbox'
    'LAPS' = 'LAPS'
    'LAPS-Ablauf' = 'LAPS expiry'
    'Laufwerk' = 'Drive'
    'Lease-Dauer' = 'Lease duration'
    'Leitet weiter an' = 'Forwards to'
    'Leseweg' = 'Read path'
    'Letzte Aktualisierung' = 'Last update'
    'Letzte Anmeldung' = 'Last logon'
    'Letzte Vollsicherung' = 'Last full backup'
    'Letzte inkrementelle Sicherung' = 'Last incremental backup'
    'Letzter' = 'Last'
    'Letzter Aufräumlauf' = 'Last cleanup run'
    'Letzter Start' = 'Last start'
    'Lokale Admins' = 'Local admins'
    'Mindestlänge' = 'Minimum length'
    'Mitglied' = 'Member'
    'Mitglieder' = 'Members'
    'NLA' = 'NLA'
    'NTLM' = 'NTLM'
    'Name' = 'Name'
    'Nr.' = 'No.'
    'Objekt' = 'Object'
    'Ordner' = 'Folder'
    'Organisationseinheit' = 'OU'
    'Ort' = 'Location'
    'PDC-Emulator' = 'PDC emulator'
    'Pfad' = 'Path'
    'Postfach' = 'Mailbox'
    'Postfach-Typ' = 'Mailbox type'
    'Postfachtyp' = 'Mailbox type'
    'Postfächer' = 'Mailboxes'
    'Privilegierte Gruppen' = 'Privileged groups'
    'Programm' = 'Program'
    'Prüfpunkt' = 'Checkpoint'
    'Quelle' = 'Source'
    'Quelle (Anzahl)' = 'Sources (count)'
    'RID-Master' = 'RID master'
    'Recht' = 'Permission'
    'Rechte' = 'Permissions'
    'Regel' = 'Rule'
    'Reihenfolge' = 'Order'
    'SPN' = 'SPN'
    'Schemamaster' = 'Schema master'
    'Schlüssel' = 'Key'
    'Schreibgeschützt' = 'Read-only'
    'Sendesperre ab' = 'Send blocked from'
    'Server' = 'Server'
    'Sid' = 'SID'
    'Signatur' = 'Signature'
    'Speicher' = 'Storage'
    'Sperrungen' = 'Lockouts'
    'Stand' = 'Status'
    'Standort' = 'Site'
    'Status' = 'Status'
    'Support-Ende' = 'End of support'
    'Support-Ende (spätestes)' = 'End of support (latest)'
    'Tage bis Ablauf' = 'Days until expiry'
    'Tage her' = 'Days ago'
    'Tage ohne Anmeldung' = 'Days without logon'
    'Tage seit Ablauf' = 'Days since expiry'
    'Treffer' = 'Hits'
    'Typ' = 'Type'
    'Vererbung' = 'Inheritance'
    'Verschiedene Konten' = 'Distinct accounts'
    'Verschlüsselt' = 'Encrypted'
    'Version' = 'Version'
    'Version C/B' = 'Version C/B'
    'Verweigert für' = 'Denied for'
    'Verwendung' = 'Use'
    'Virenschutz' = 'Antivirus'
    'Von' = 'From'
    'WMI-Filter' = 'WMI filter'
    'Weg' = 'Path'
    'Weiterleitungen' = 'Forwarding'
    'Wert' = 'Value'
    'Wirksame Richtlinie' = 'Effective policy'
    'Wirksames Recht' = 'Effective permission'
    'Wirkung' = 'Effect'
    'Zeit' = 'Time'
    'Zeitpunkt' = 'Time'
    'Zeitraum' = 'Period'
    'Zertifikat' = 'Certificate'
    'Zone' = 'Zone'
    'Zonen' = 'Zones'
    'Zonenübertragung' = 'Zone transfer'
    'Zugriff für große Gruppen' = 'Access for large groups'
    'Zugriffsbasierte Aufzählung' = 'Access-based enumeration'
    'Zuletzt' = 'Last'
    'Zustand' = 'State'
    'davon Größe (GB)' = 'Of which size (GB)'
    'davon aktiv' = 'Of which active'
    'davon inaktiv' = 'Of which inactive'
    'davon unbekannte Konten' = 'Of which unknown accounts'
    'ist Mitglied von' = 'Member of'
    'Älter als 5 Jahre' = 'Older than 5 years'
    'Älter als 5 Jahren' = 'Older than 5 years'
    'Ältester Eintrag' = 'Oldest entry'
    'Über' = 'Via'
    'Überwachungsprotokoll' = 'Audit log'
}
# Werte, die der Code intern auf Deutsch vergleicht (z. B. Art = 'Benutzer'), erscheinen im englischen Bericht übersetzt (nur Anzeige und CSV)
$WerteEn = New-Object System.Collections.Specialized.OrderedDictionary([System.StringComparer]::Ordinal)   # Groß-/Kleinschreibung zählt (Verwaist / verwaist)
$WerteEn.Add('Alle erweiterten Rechte (u. a. Kennwort zurücksetzen)', 'All extended rights (including password reset)')
$WerteEn.Add('Kein Serverbetriebssystem (Support nicht bewertet)', 'No server operating system (support not assessed)')
$WerteEn.Add('Rechte je Ordner (nur dort, wo sie gesetzt werden)', 'Permissions per folder (only where set)')
$WerteEn.Add('Mehrfach freigegebene Ordner (nur einmal gezählt)', 'Folders shared multiple times (counted only once)')
$WerteEn.Add('Übersprungene Verknüpfungen (Junctions/Symlinks)', 'Skipped links (Junctions/Symlinks)')
$WerteEn.Add('Der Ordner der Freigabe existiert nicht mehr.', 'The share folder no longer exists.')
$WerteEn.Add('DHCP-Überwachungsprotokoll ist ausgeschaltet', 'DHCP audit log is disabled')
$WerteEn.Add('Vollzugriff und "Senden als" im Überblick', 'Full access and "Send as" overview')
$WerteEn.Add('Ordner der Freigabe existiert nicht mehr', 'Share folder no longer exists')
$WerteEn.Add('Enthält außer Sicherheitseinstellungen:', 'Contains except security settings:')
$WerteEn.Add('Nicht geprüft - Postfach nicht gefunden', 'Not checked - mailbox not found')
$WerteEn.Add('WMI-Filter nicht lesbar - nicht geprüft', 'WMI filter not readable - not checked')
$WerteEn.Add('Abfrage fehlgeschlagen - nicht geprüft', 'Query failed - not checked')
$WerteEn.Add('Nicht im Active Directory gespeichert', 'Not stored in Active Directory')
$WerteEn.Add('keine im Active Directory autorisiert', 'None authorized in Active Directory')
$WerteEn.Add('Zugriff auf den Ordner nicht möglich', 'Access to folder not possible')
$WerteEn.Add('Zertifikate (öffentliche Schlüssel)', 'Certificates (public keys)')
$WerteEn.Add('Domänencontroller der Organisation', 'Organization domain controllers')
$WerteEn.Add('Remotedesktop erlaubt Verbindungen', 'Remote Desktop allows connections')
$WerteEn.Add('autorisierte DHCP-Server ermitteln', 'Determine authorized DHCP servers')
$WerteEn.Add('Windows Server 2008 R2 und älter', 'Windows Server 2008 R2 and older')
$WerteEn.Add('Konto deaktiviert oder gesperrt', 'Account disabled or locked')
$WerteEn.Add('Vollzugriff (Zusatz: ObjectType', 'Full access (additional: ObjectType')
$WerteEn.Add('Kennwort oder Konto abgelaufen', 'Password or account expired')
$WerteEn.Add('Lokale Administratorengruppen', 'Local administrator groups')
$WerteEn.Add('Berechtigungen/Besitz ändern', 'Change permissions/ownership')
$WerteEn.Add('Freigabepfade nicht sichtbar', 'Share paths not visible')
$WerteEn.Add('Aktive privilegierte Konten', 'Active privileged accounts')
$WerteEn.Add('Gruppe hat keine Mitglieder', 'Group has no members')
$WerteEn.Add('Keine Verbindung zum Server', 'No connection to server')
$WerteEn.Add('unerlaubte Zeichen im Namen', 'Illegal characters in name')
$WerteEn.Add('Geprüfte Domänencontroller', 'Checked domain controllers')
$WerteEn.Add('Authentifizierte Benutzer', 'Authenticated Users')
$WerteEn.Add('Schwellen dieses Berichts', 'Thresholds of this report')
$WerteEn.Add('Benutzer (inetOrgPerson)', 'User (inetOrgPerson)')
$WerteEn.Add('Datenbestand je Freigabe', 'Data inventory per share')
$WerteEn.Add('Kontosperrungsrichtlinie', 'Account lockout policy')
$WerteEn.Add('Sicherheitseinstellungen', 'Security settings')
$WerteEn.Add('Alle erweiterten Rechte', 'All extended rights')
$WerteEn.Add('Hinweise zur Auswertung', 'Notes for evaluation')
$WerteEn.Add('Nicht abgefragte Server', 'Servers not queried')
$WerteEn.Add('Unterordner und Dateien', 'Subfolders and files')
$WerteEn.Add('Bestätigtes Schreiben:', 'Validated write:')
$WerteEn.Add('Installierte Programme', 'Installed programs')
$WerteEn.Add('Konto für DNS-Einträge', 'Account for DNS records')
$WerteEn.Add('Kontosperrungsschwelle', 'Account lockout threshold')
$WerteEn.Add('Nicht lesbare Bereiche', 'Unreadable areas')
$WerteEn.Add('Berechtigungen ändern', 'Change permissions')
$WerteEn.Add('Hinweise zum Auslesen', 'Notes for reading')
$WerteEn.Add('Windows 8.1 und älter', 'Windows 8.1 and older')
$WerteEn.Add('Überwachungsprotokoll', 'Audit log')
$WerteEn.Add('Nicht berücksichtigt', 'Not considered')
$WerteEn.Add('Anonymous-Anmeldung', 'Anonymous logon')
$WerteEn.Add('Infrastrukturmaster', 'Infrastructure master')
$WerteEn.Add('Secure Boot ist aus', 'Secure Boot is off')
$WerteEn.Add('Domänennamenmaster', 'Domain naming master')
$WerteEn.Add('Kennwortrichtlinie', 'Password policy')
$WerteEn.Add('Komponentenwartung', 'Component servicing')
$WerteEn.Add('alle Eigenschaften', 'all properties')
$WerteEn.Add('Besitz übernehmen', 'Take ownership')
$WerteEn.Add('Domänencontroller', 'Domain Controllers')
$WerteEn.Add('Falsches Kennwort', 'Incorrect password')
$WerteEn.Add('Konten-Operatoren', 'Account Operators')
$WerteEn.Add('Nur dieser Ordner', 'Only this folder')
$WerteEn.Add('Domänen-Benutzer', 'Domain Users')
$WerteEn.Add('Format unbekannt', 'Format unknown')
$WerteEn.Add('Sonstige Dateien', 'Other files')
$WerteEn.Add('Ältester Eintrag', 'Oldest entry')
$WerteEn.Add('Domänencomputer', 'Domain Computers')
$WerteEn.Add('Konto unbekannt', 'Account unknown')
$WerteEn.Add('Standard (alle)', 'Default (all)')
$WerteEn.Add('Computerkonten', 'Computer accounts')
$WerteEn.Add('Konto gesperrt', 'Account locked')
$WerteEn.Add('Nicht erreicht', 'Not reached')
$WerteEn.Add('DHCP-Bereiche', 'DHCP scopes')
$WerteEn.Add('Dieser Ordner', 'This folder')
$WerteEn.Add('Gruppe/System', 'Group/system')
$WerteEn.Add('Remotedesktop', 'Remote Desktop')
$WerteEn.Add('Arbeitsplatz', 'Workstation')
$WerteEn.Add('Schemamaster', 'Schema master')
$WerteEn.Add('DHCP-Server', 'DHCP server')
$WerteEn.Add('Datenbanken', 'Databases')
$WerteEn.Add('Sonderkonto', 'Special account')
$WerteEn.Add('Unterordner', 'Subfolders')
$WerteEn.Add('Vollzugriff', 'Full control')
$WerteEn.Add('Zugriff von', 'Access from')
$WerteEn.Add('DNS-Server', 'DNS server')
$WerteEn.Add('Inaktiv ab', 'Inactive from')
$WerteEn.Add('Schreiben:', 'Write:')
$WerteEn.Add('Verweigern', 'Deny')
$WerteEn.Add('DNS-Zonen', 'DNS zones')
$WerteEn.Add('Schreiben', 'Write')
$WerteEn.Add('Sonstiges', 'Other')
$WerteEn.Add('Standorte', 'Sites')
$WerteEn.Add('Unbekannt', 'Unknown')
$WerteEn.Add('unbekannt', 'unknown')
$WerteEn.Add('(Wurzel)', '(root)')
$WerteEn.Add('Anlegen:', 'Create:')
$WerteEn.Add('Benutzer', 'User')
$WerteEn.Add('Erlauben', 'Allow')
$WerteEn.Add('Netzwerk', 'Network')
$WerteEn.Add('Verwaist', 'Orphaned')
$WerteEn.Add('verwaist', 'orphaned')
$WerteEn.Add('Dateien', 'Files')
$WerteEn.Add('Gruppen', 'Groups')
$WerteEn.Add('Anonym', 'Anonymous')
$WerteEn.Add('Domäne', 'Domain')
$WerteEn.Add('Gruppe', 'Group')
$WerteEn.Add('Umfang', 'Scope')
$WerteEn.Add('direkt', 'direct')
$WerteEn.Add('Ändern', 'Change')
$WerteEn.Add('Lesen', 'Read')
$WerteEn.Add('über', 'via')
$WerteMuster = @()
foreach ($k in ($WerteEn.Keys | Sort-Object Length -Descending)) { $WerteMuster += , @([regex]::new('(?<![\p{L}\d])' + [regex]::Escape($k) + '(?![\p{L}\d])'), $WerteEn[$k]) }
function WertAnzeige($v) {
    if (-not $script:En -or $v -isnot [string] -or -not $v) { return $v }
    if ($v.Contains('\')) { return $v }   # Kontonamen (DOMÄNE\Name) und Pfade bleiben, wie sie sind
    foreach ($m in $WerteMuster) { $v = $m[0].Replace($v, $m[1]) }
    $v
}
function SpaltenAnzeige($liste) {
    if (-not $script:En -or -not @($liste).Count) { return $liste }
    foreach ($o in @($liste)) {
        $n = [ordered]@{}
        foreach ($p in $o.PSObject.Properties) { $k = $p.Name; if ($SpaltenEn.ContainsKey($k)) { $k = $SpaltenEn[$k] }; $i = 2; $k0 = $k; while ($n.Contains($k)) { $k = "$k0 $i"; $i++ }; $n[$k] = WertAnzeige $p.Value }
        [pscustomobject]$n
    }
}
# Sortierschluessel fuer Nummern wie K2, K10: erst Buchstaben, dann die Zahl als Zahl
function IdZahl($id) { $m = [regex]::Match([string]$id, '\d+'); if ($m.Success) { [int]$m.Value } else { 0 } }
function IdText($id) { [regex]::Replace([string]$id, '\d+', '') }

$Befunde = New-Object System.Collections.Generic.List[object]

# Platzhalterzeilen ("nicht geprüft") erkennt das Skript nur an den vollständigen, festen Texten, die es selbst erzeugt - keine Teilstring-Suche
# in Konto-, Objekt- oder Beschreibungstexten, die ein Angreifer beeinflussen könnte
# Deutsch und Englisch: Skripte erzeugen den Text je nach -Language, erkannt werden beide Fassungen
$PlatzhalterTexte = @('Abfrage fehlgeschlagen - nicht geprüft', 'Berechtigungen nicht lesbar - nicht geprüft', 'SYSVOL nicht lesbar - nicht geprüft', 'WMI-Filter nicht lesbar - nicht geprüft', 'Kein Domänencontroller ermittelbar - SYSVOL nicht geprüft', 'Inhalt des Ordners nicht lesbar - Richtlinie nicht geprüft',
    'Query failed - not checked', 'Permissions not readable - not checked', 'SYSVOL not readable - not checked', 'WMI filter not readable - not checked', 'No domain controller found - SYSVOL not checked', 'Folder content not readable - policy not checked')
function IstPlatzhalter($o) {
    # Ein Skript kann $script:PlatzhalterSpalten setzen: dann zählen nur diese Spalten (Namensspalten sind von Benutzern beeinflussbar)
    foreach ($pr in $o.PSObject.Properties) {
        if ($script:PlatzhalterSpalten -and ($pr.Name -notin $script:PlatzhalterSpalten)) { continue }
        if (($pr.Value -is [string]) -and ($PlatzhalterTexte -ccontains $pr.Value)) { return $true }
    }
    ("$($o.Feststellung)" -clike 'Nicht geprüft - *') -or ("$($o.Feststellung)" -clike 'Not checked - *')
}

# Besteht ein Prüfpunkt nur aus Platzhalterzeilen (Abfrage fehlgeschlagen), ist das kein Mangel dieses Prüfpunkts: als Info führen statt
# mit seiner Stufe, damit es in den Kacheln nicht als hoher Befund zählt. Nur für Prüfpunkte, deren Ergebnis von einer Abfrage abhängt.
function InfoWennNurNichtGeprueft([string]$IdMuster) {
    foreach ($x in ($Befunde | Where-Object { $_.Id -match $IdMuster -and $_.Anzahl -gt 0 -and $_.Stufe -ne 'Info' })) {
        if (-not @($x.Objekte | Where-Object { -not (IstPlatzhalter $_) }).Count) { $x.Stufe = 'Info' }
    }
}

# Prüfpunkte, die für einzelne Server nicht ausgewertet werden konnten (Abfrage fehlgeschlagen), dürfen nicht als "ohne Befund" erscheinen:
# je Server eine Zeile "nicht geprüft" anhängen; hatte der Prüfpunkt sonst keinen Treffer, steht er als Info statt als OK.
function NichtGeprueft([string]$IdMuster, [string[]]$Namen, [string]$Verweis, [string]$Spalte = 'Server', [string]$Grund = $(T 'Abfrage nicht möglich' 'query not possible')) {
    foreach ($x in ($Befunde | Where-Object { $_.Id -match $IdMuster })) {
        $war = $x.Anzahl
        $zeilen = @($Namen | ForEach-Object { [pscustomobject][ordered]@{ $Spalte = $_; Feststellung = (T "Nicht geprüft - $Grund (siehe $Verweis)" "Not checked - $Grund (see $Verweis)"); Details = '' } })
        $x.Objekte = @($x.Objekte) + $zeilen
        $x.Anzahl = $x.Objekte.Count
        if (-not $war) { $x.Stufe = 'Info' }
    }
}
function Befund {
    param([string]$Id, [string]$Bereich, [string]$Titel, [ValidateSet('Hoch', 'Mittel', 'Niedrig', 'Info')][string]$Stufe,
        $Objekte, [string]$Erklaerung, [string]$Empfehlung)
    $liste = @($Objekte | Where-Object { $null -ne $_ })
    $status = $Stufe
    if ($liste.Count -eq 0) { $status = 'OK' }
    $Befunde.Add([pscustomobject]@{
            Id = $Id; Bereich = $Bereich; Titel = $Titel; Stufe = $status; Anzahl = $liste.Count
            Objekte = $liste; Erklaerung = $Erklaerung; Empfehlung = $Empfehlung
        })
}

# Schreibt den HTML-Bericht (und auf Wunsch CSV-Dateien). $Abschnitte: Überschrift -> Objekte oder fertiges HTML,
# erscheinen vor den Befunden. $Anhang: Überschrift -> Text, erscheint am Ende unverändert.
function BerichtSchreiben {
    param([string]$Name, [string]$Datei, [string]$Bezug, [string]$BezugName = $(T 'Domäne' 'Domain'), [System.Collections.IDictionary]$Umgebung,
        [System.Collections.IDictionary]$Abschnitte = @{}, [System.Collections.IDictionary]$Anhang = @{}, [string]$Fuss)
    Write-Host (T '  Bericht schreiben ...' '  Writing report ...')
    $jetzt = Get-Date
    $ziel = $script:Ausgabe
    if (-not $ziel) { $ziel = Join-Path (Get-Location).Path ("{0}_{1}_{2:yyyy-MM-dd_HHmm}.html" -f $Datei, ($Bezug -replace '[^\w.\-]+', '-'), $jetzt) }
    $ziel = [System.IO.Path]::GetFullPath($ziel)
    $ordner = Split-Path $ziel -Parent
    if (-not (Test-Path $ordner)) { New-Item -ItemType Directory -Path $ordner -Force | Out-Null }

    $stufeText = @{ Hoch = (T 'Hoch' 'High'); Mittel = (T 'Mittel' 'Medium'); Niedrig = (T 'Niedrig' 'Low'); Info = 'Info'; OK = 'OK' }
    $rang = @{ Hoch = 0; Mittel = 1; Niedrig = 2; Info = 3; OK = 4 }
    $zahl = @{}
    foreach ($s in 'Hoch', 'Mittel', 'Niedrig', 'Info', 'OK') { $zahl[$s] = @($Befunde | Where-Object Stufe -eq $s).Count }
    $css = @'
body{font-family:Segoe UI,Arial,sans-serif;color:#2b2f36;margin:0;background:#f4f6fa}
.seite{max-width:1100px;margin:0 auto;padding:28px 24px 60px}
h1{color:#1f5fae;margin:0 0 4px;font-size:28px} h2{color:#1f5fae;font-size:19px;margin:34px 0 10px;border-bottom:2px solid #d0d5dd;padding-bottom:6px}
.meta{color:#6b7280;font-size:13px;margin-bottom:22px}
.kacheln{display:flex;gap:14px;flex-wrap:wrap;margin:18px 0}
.kachel{flex:1;min-width:150px;background:#fff;border-radius:10px;padding:14px 18px;border-left:8px solid #999;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.kachel b{display:block;font-size:30px} .kachel span{font-size:13px;color:#6b7280}
table{border-collapse:collapse;width:100%;background:#fff;font-size:13px;margin:8px 0 4px}
th{background:#1f5fae;color:#fff;text-align:left;padding:7px 9px;font-weight:600} td{padding:6px 9px;border-bottom:1px solid #e5e7eb;vertical-align:top;overflow-wrap:break-word}
tr:nth-child(even) td{background:#f9fafb}
.marke{display:inline-block;min-width:58px;text-align:center;padding:2px 9px;border-radius:11px;color:#fff;font-size:12px;font-weight:600}
.m-Hoch{background:#d92d20}.m-Mittel{background:#f79009}.m-Niedrig{background:#ca8a04}.m-OK{background:#12b76a}.m-Info{background:#1f5fae}
.befund{background:#fff;border-radius:10px;padding:16px 20px;margin:14px 0;border-left:8px solid #999;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.befund h3{margin:0 0 6px;font-size:16px} .befund p{margin:5px 0;font-size:14px;line-height:1.45} .befund .emp{color:#14532d}
.Hoch{border-color:#d92d20}.Mittel{border-color:#f79009}.Niedrig{border-color:#eab308}.OK{border-color:#12b76a}.Info{border-color:#1f5fae}
.klein{color:#6b7280;font-size:12px} a{color:#1f5fae;text-decoration:none}
pre{background:#fff;border:1px solid #d0d5dd;border-radius:8px;padding:12px 14px;font-size:12px;white-space:pre-wrap;word-break:break-word}
.fuss{margin-top:40px;color:#6b7280;font-size:12px;border-top:1px solid #d0d5dd;padding-top:12px}
@media print{body{background:#fff}.befund,.kachel{box-shadow:none;border:1px solid #d0d5dd;border-left-width:8px;break-inside:avoid}}
'@
    $h = New-Object System.Text.StringBuilder
    [void]$h.Append("<!DOCTYPE html><html lang='$Language'><head><meta charset='utf-8'><title>$(Html $Name) $(Html $Bezug)</title><style>$css</style></head><body><div class='seite'>")
    [void]$h.Append("<h1>$(Html $Name)</h1><div class='meta'>$(Html $BezugName) <b>$(Html $Bezug)</b> · $(T 'erstellt am' 'created') $($jetzt.ToString((T 'dd.MM.yyyy HH:mm' 'yyyy-MM-dd HH:mm'))) $(T 'von' 'by') $(Html "$env:USERDOMAIN\$env:USERNAME") $(T 'auf' 'on') $(Html $env:COMPUTERNAME)</div>")
    [void]$h.Append("<div class='kacheln'>")
    foreach ($s in 'Hoch', 'Mittel', 'Niedrig', 'Info', 'OK') {
        $text = @{ Hoch = (T 'Befunde mit hoher Priorität' 'findings, high priority'); Mittel = (T 'mit mittlerer Priorität' 'medium priority'); Niedrig = (T 'mit niedriger Priorität' 'low priority'); Info = (T 'Hinweise und nicht geprüft' 'notes and not checked'); OK = (T 'Prüfpunkte ohne Befund' 'checks without findings') }[$s]
        [void]$h.Append("<div class='kachel $s'><b>$($zahl[$s])</b><span>$text</span></div>")
    }
    [void]$h.Append("</div><h2>$(T 'Umgebung' 'Environment')</h2><table><tr><th>$(T 'Eigenschaft' 'Property')</th><th>$(T 'Wert' 'Value')</th></tr>")
    foreach ($k in $Umgebung.Keys) { [void]$h.Append("<tr><td>$(Html (WertAnzeige $k))</td><td>$(Html (WertAnzeige $Umgebung[$k]))</td></tr>") }
    [void]$h.Append("</table>")

    [void]$h.Append("<h2>$(T 'Übersicht der Prüfpunkte' 'Overview of checks')</h2><table><tr><th>$(T 'Nr.' 'No.')</th><th>$(T 'Bereich' 'Area')</th><th>$(T 'Prüfpunkt' 'Check')</th><th>$(T 'Treffer' 'Hits')</th><th>$(T 'Bewertung' 'Rating')</th></tr>")
    foreach ($x in $Befunde) {
        [void]$h.Append("<tr><td>$($x.Id)</td><td>$(Html $x.Bereich)</td><td><a href='#$($x.Id)'>$(Html $x.Titel)</a></td><td>$($x.Anzahl)</td><td><span class='marke m-$($x.Stufe)'>$($stufeText[$x.Stufe])</span></td></tr>")
    }
    [void]$h.Append("</table>")

    foreach ($k in $Abschnitte.Keys) {
        [void]$h.Append("<h2>$(Html (WertAnzeige $k))</h2>")
        $inhalt = $Abschnitte[$k]
        if ($inhalt -is [string]) { [void]$h.Append($inhalt) }
        elseif (@($inhalt).Count) { [void]$h.Append(((Einheitlich $inhalt | ConvertTo-Html -Fragment) -join "`n")) }
        else { [void]$h.Append("<p class='klein'>$(T 'Keine Einträge.' 'No entries.')</p>") }
    }

    [void]$h.Append("<h2>$(T 'Befunde' 'Findings')</h2>")
    $offen = @($Befunde | Where-Object { $_.Stufe -ne 'OK' } | Sort-Object { $rang[$_.Stufe] }, { IdText $_.Id }, { IdZahl $_.Id })
    foreach ($x in $offen) {
        [void]$h.Append("<div class='befund $($x.Stufe)' id='$($x.Id)'><h3><span class='marke m-$($x.Stufe)'>$($stufeText[$x.Stufe])</span> $($x.Id) · $(Html $x.Titel) <span class='klein'>($($x.Anzahl) $(T 'Treffer' 'hits'))</span></h3>")
        [void]$h.Append("<p>$(Html $x.Erklaerung)</p><p class='emp'><b>$(T 'Empfehlung' 'Recommendation'):</b> $(Html $x.Empfehlung)</p>")
        [void]$h.Append(((Einheitlich $x.Objekte | ConvertTo-Html -Fragment) -join "`n"))
        [void]$h.Append("</div>")
    }
    if (-not $offen.Count) { [void]$h.Append("<p>$(T 'Keine Auffälligkeiten gefunden.' 'No anomalies found.')</p>") }
    $ok = @($Befunde | Where-Object Stufe -eq 'OK')
    if ($ok.Count) {
        [void]$h.Append("<h2>$(T 'Ohne Befund' 'No findings')</h2><table><tr><th>$(T 'Nr.' 'No.')</th><th>$(T 'Prüfpunkt' 'Check')</th></tr>")
        foreach ($x in $ok) { [void]$h.Append("<tr id='$($x.Id)'><td>$($x.Id)</td><td>$(Html $x.Titel)</td></tr>") }
        [void]$h.Append("</table>")
    }
    foreach ($k in $Anhang.Keys) { [void]$h.Append("<h2>$(Html $k)</h2><pre>$(Html $Anhang[$k])</pre>") }
    [void]$h.Append("<div class='fuss'>$(Html $Name) Version $script:Version · IT Admin Tools by MM · $(T 'Das Skript hat ausschließlich gelesen und nichts verändert.' 'The script only read and changed nothing.') $(Html $Fuss)</div>")
    [void]$h.Append("</div></body></html>")
    [System.IO.File]::WriteAllText($ziel, $h.ToString(), (New-Object System.Text.UTF8Encoding($true)))

    $basis = [System.IO.Path]::ChangeExtension($ziel, $null).TrimEnd('.')
    if ($script:Csv) {
        foreach ($x in ($Befunde | Where-Object { $_.Anzahl -gt 0 })) {
            Einheitlich $x.Objekte | Export-Csv -Path "${basis}_$($x.Id).csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8
        }
    }
    Write-Host ''
    Write-Host ((T "Fertig: {0} hoch, {1} mittel, {2} niedrig, {3} ohne Befund." "Done: {0} high, {1} medium, {2} low, {3} without findings.") -f $zahl.Hoch, $zahl.Mittel, $zahl.Niedrig, $zahl.OK) -ForegroundColor Green
    Write-Host ((T 'Bericht' 'Report') + ": $ziel")
    if (-not $script:NichtOeffnen -and [Environment]::UserInteractive) { try { Start-Process $ziel } catch { } }
    $basis
}

# ------------------------------------------------------------------ Verbindung
$session = $null
trap { if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }; break }
function ExchangeBereit { [bool](Get-Command Get-MailboxFolderPermission -ErrorAction SilentlyContinue) }
if (-not (ExchangeBereit) -and -not $ExchangeServer) {
    $kandidaten = @()
    if ($env:ExchangeInstallPath) { $kandidaten += (Join-Path $env:ExchangeInstallPath 'bin\RemoteExchange.ps1') }
    $kandidaten += 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'
    $rex = $kandidaten | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if ($rex) {
        Write-Host (T 'Lade Exchange-Verwaltungsshell ...' 'Loading Exchange Management Shell ...') -ForegroundColor Cyan
        $ErrorActionPreference = 'Continue'
        try { . $rex *> $null; Connect-ExchangeServer -auto -ClientApplication:ManagementShell *> $null }
        catch { Write-Warning (T "Die Exchange-Verwaltungsshell ließ sich nicht laden: $($_.Exception.Message)" "The Exchange Management Shell could not be loaded: $($_.Exception.Message)") }
        $ErrorActionPreference = 'Stop'
    }
}
if (-not (ExchangeBereit)) {
    if (-not $ExchangeServer) { throw (T 'Die Exchange-Befehle sind nicht verfügbar. Bitte auf einem Exchange-Server starten oder -ExchangeServer angeben.' 'The Exchange cmdlets are not available. Please start on an Exchange server or specify -ExchangeServer.') }
    Write-Host (T "Baue Remote-Sitzung zu $ExchangeServer auf ..." "Building remote session to $ExchangeServer ...") -ForegroundColor Cyan
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$ExchangeServer/PowerShell/" -Authentication Kerberos -SessionOption (New-PSSessionOption -OpenTimeout 60000)
    Import-PSSession -Session $session -DisableNameChecking -AllowClobber -CommandName Get-Mailbox, Get-MailboxFolderStatistics, Get-MailboxFolderPermission 3>$null | Out-Null
}

# ------------------------------------------------------------------ Lesen
Write-Host (T 'Kalender-Check - es wird nur gelesen, nichts verändert.' 'Calendar check - only reading is performed, nothing is changed.') -ForegroundColor Cyan
$mbx = @()
$nichtGefunden = New-Object System.Collections.Generic.List[object]
if ($nurAuswahl) {
    foreach ($id in $Postfach) {
        try { $mbx += @(Get-Mailbox -Identity $id -ErrorAction Stop) }
        catch { $nichtGefunden.Add([pscustomobject][ordered]@{ Postfach = $id; Feststellung = 'Nicht geprüft - Postfach nicht gefunden'; Details = Kurz $_.Exception.Message 160 }) }
    }
}
else { $mbx = @(Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox) }
$mbx = @($mbx | Sort-Object DisplayName)
Write-Host (T "$($mbx.Count) Postfächer werden gelesen ..." "Reading $($mbx.Count) mailboxes ...") -ForegroundColor Cyan

$rechte = New-Object System.Collections.Generic.List[object]
$nichtGeprueft = New-Object System.Collections.Generic.List[object]
foreach ($x in $nichtGefunden) { $nichtGeprueft.Add($x) }
$nr = 0
foreach ($m in $mbx) {
    $nr++
    Write-Progress -Activity 'Lese Kalenderrechte' -Status "$nr von $($mbx.Count): $($m.DisplayName)" -PercentComplete ([int](100 * $nr / [math]::Max(1, $mbx.Count)))
    $smtp = [string]$m.PrimarySmtpAddress
    try {
        $f = Get-MailboxFolderStatistics -Identity $smtp -FolderScope Calendar -ErrorAction Stop | Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object -First 1
        if (-not $f) { throw (T 'Kein Standardkalender gefunden (Postfach evtl. nie angemeldet).' 'No default calendar found (mailbox possibly never logged on).') }
        $pfad = "${smtp}:" + ([string]$f.FolderPath -replace '/', '\')
        foreach ($p in @(Get-MailboxFolderPermission -Identity $pfad -ErrorAction Stop)) {
            $typ = ''; try { $typ = [string]$p.User.UserType } catch { }
            $art = if ($typ -match '^(Default|Anonymous|Unknown|External)$') { $typ } else { 'Intern' }
            if ($art -eq 'Intern' -and [string]$p.User -match '^(Default|Standard)$') { $art = 'Default' }
            if ($art -eq 'Intern' -and [string]$p.User -match '^(Anonymous|Anonym)$') { $art = 'Anonymous' }
            $gruppe = $false; try { $gruppe = ([string]$p.User.ADRecipient.RecipientTypeDetails) -match 'Group|RoomList' } catch { }
            $rechte.Add([pscustomobject]@{
                    Postfach = [string]$m.DisplayName; Adresse = $smtp; Typ = [string]$m.RecipientTypeDetails
                    Berechtigter = $(if ($p.User.DisplayName) { [string]$p.User.DisplayName } else { [string]$p.User }); Art = $art; Gruppe = $gruppe
                    Recht = (@($p.AccessRights) -join ', ')
                })
        }
    }
    catch { $nichtGeprueft.Add([pscustomobject][ordered]@{ Postfach = $smtp; Feststellung = (T 'Nicht geprüft - Kalenderrechte nicht lesbar' 'Not checked - calendar permissions not readable'); Details = Kurz $_.Exception.Message 160 }) }
}
Write-Progress -Activity 'Lese Kalenderrechte' -Completed
if ($session) { Remove-PSSession $session }

# ------------------------------------------------------------------ Prüfungen
$schreib = 'Author', 'PublishingAuthor', 'NonEditingAuthor', 'Contributor', 'Editor', 'PublishingEditor', 'Owner', 'FolderOwner'
function IstSchreibrecht([string]$r) { foreach ($x in ($r -split ',\s*')) { if ($x -in $schreib -or $x -match '^(CreateItems|EditAllItems|EditOwnedItems|DeleteAllItems|DeleteOwnedItems)$') { return $true } }; $false }
function PostfachTyp([string]$t) { switch ($t) { 'UserMailbox' { (T 'Benutzer' 'User') } 'SharedMailbox' { 'Gemeinsam' } 'RoomMailbox' { 'Raum' } 'EquipmentMailbox' { (T 'Gerät' 'Device') } default { $t } } }
function Zeile($r, $zusatz = @{}) {
    $o = [ordered]@{ Postfach = $r.Postfach; Adresse = $r.Adresse; 'Postfach-Typ' = (PostfachTyp $r.Typ); Berechtigter = $r.Berechtigter; Recht = $r.Recht }
    foreach ($k in $zusatz.Keys) { $o[$k] = $zusatz[$k] }
    [pscustomobject]$o
}
$nurFrei = 'None', 'AvailabilityOnly'

Befund 'K1' (T 'Standardzugriff' 'Default access') (T 'Alle Benutzer dürfen mehr als Frei/Gebucht sehen (Standard-Recht)' 'All users are allowed to see more than Free/Busy (default permission)') 'Mittel' `
    @($rechte | Where-Object { $_.Art -eq 'Default' -and $_.Recht -notin $nurFrei } | Sort-Object Postfach | ForEach-Object { Zeile $_ }) `
    (T 'Das Standard-Recht gilt für jeden Benutzer der Organisation ohne eigenen Eintrag. Ab "LimitedDetails" sehen alle Betreff und Ort, ab "Reviewer" den ganzen Termin. Bei Kalendern von Geschäftsführung, Personal oder Betriebsrat ist das selten gewollt.' 'The default permission applies to every user in the organization without a separate entry. At "LimitedDetails", all users see subject and location; at "Reviewer", they see the entire meeting. For calendars of management, HR, or works council, this is rarely desired.') `
    (T 'Standard-Recht auf "AvailabilityOnly" (nur Frei/Gebucht) setzen, wo mehr nicht nötig ist, und gezielt Berechtigte eintragen.' 'Set the default permission to "AvailabilityOnly" (free/busy only) where more is not needed, and add specific authorized users.')
Befund 'K2' (T 'Standardzugriff' 'Default access') (T 'Anonymer Zugriff auf einen Kalender' 'Anonymous access to a calendar') 'Hoch' `
    @($rechte | Where-Object { $_.Art -eq 'Anonymous' -and $_.Recht -ne 'None' } | Sort-Object Postfach | ForEach-Object { Zeile $_ }) `
    (T 'Anonymer Zugriff erlaubt Personen ohne Anmeldung, den Kalender zu sehen, zum Beispiel über eine veröffentlichte Adresse im Internet.' 'Anonymous access allows users without logon to view the calendar, for example via a published address on the internet.') `
    (T 'Anonym auf "None" setzen, wenn nicht ausdrücklich eine Veröffentlichung gewollt ist.' 'Set anonymous to "None" if publication is not explicitly desired.')
Befund 'K3' (T 'Schreibrechte' 'Write permissions') (T 'Personen oder Gruppen mit Schreibrecht auf einen fremden Kalender' 'Persons or groups with write permission to another user''s calendar') 'Info' `
    @($rechte | Where-Object { $_.Art -in 'Intern', 'External' -and (IstSchreibrecht $_.Recht) } | Sort-Object Postfach, Berechtigter | ForEach-Object { Zeile $_ @{ Art = $(if ($_.Gruppe) { (T 'Gruppe' 'Group') } else { 'Person' }) } }) `
    (T 'Mit Schreibrecht können Termine angelegt, geändert und (je nach Recht) gelöscht werden. Das ist bei Stellvertretern und Sekretariaten üblich, sollte aber bekannt und gewollt sein.' 'With write permission, appointments can be created, modified, and (depending on the permission) deleted. This is common for delegates and secretaries, but should be known and intended.') `
    (T 'Liste mit den Verantwortlichen durchgehen und nicht mehr benötigte Rechte entfernen.' 'Review the list of responsible persons and remove permissions no longer needed.')
Befund 'K4' (T 'Aufräumen' 'Cleanup') (T 'Einträge für Konten, die nicht mehr auflösbar sind' 'Entries for accounts that can no longer be resolved') 'Niedrig' `
    @($rechte | Where-Object { $_.Art -eq 'Unknown' } | Sort-Object Postfach | ForEach-Object { Zeile $_ }) `
    (T 'Das Konto wurde gelöscht oder ist nicht erreichbar. Der Eintrag bleibt als Karteileiche stehen und verwirrt bei der Fehlersuche.' 'The account has been deleted or is unreachable. The entry remains as a stale entry and can confuse troubleshooting.') `
    (T 'Eintrag nach einem Blick entfernen (Remove-MailboxFolderPermission), das Skript ändert nichts.' 'Remove the entry after a review (Remove-MailboxFolderPermission); the script changes nothing.')
Befund 'K5' (T 'Abfrage' 'Query') (T 'Postfächer, deren Kalender nicht gelesen werden konnte' 'Mailboxes whose calendars could not be read') 'Info' $nichtGeprueft.ToArray() `
    (T 'Für diese Postfächer fehlt die Auswertung; sie sind in den übrigen Prüfpunkten nicht enthalten.' 'The evaluation is missing for these mailboxes; they are not included in the other checks.') `
    (T 'Meldung lesen (Rolle, Postfach nie angemeldet?) und das Skript erneut starten.' 'Read the message (role, mailbox never logged on?) and rerun the script.')

# ------------------------------------------------------------------ Bericht
$liste = $rechte | Sort-Object Postfach, Berechtigter | ForEach-Object {
    $istGruppe = $_.Gruppe   # innerhalb von switch ist $_ der geprüfte Wert, deshalb vorher merken
    [pscustomobject][ordered]@{ Postfach = $_.Postfach; Adresse = $_.Adresse; 'Postfach-Typ' = (PostfachTyp $_.Typ); Berechtigter = $_.Berechtigter; Art = $(switch ($_.Art) { 'Default' { 'Standard (alle)' } 'Anonymous' { 'Anonym' } 'Unknown' { 'unbekannt' } 'External' { 'extern' } default { if ($istGruppe) { (T 'Gruppe' 'Group') } else { 'Person' } } }); Recht = $_.Recht }
}
$umgebung = [ordered]@{
    (T 'Postfächer geprüft' 'Mailboxes checked')      = (T "$([math]::Max(0, $mbx.Count - ($nichtGeprueft.Count - $nichtGefunden.Count))) von $($mbx.Count + $nichtGefunden.Count)" "$([math]::Max(0, $mbx.Count - ($nichtGeprueft.Count - $nichtGefunden.Count))) of $($mbx.Count + $nichtGefunden.Count)")
    (T 'Einträge (gesamt)' 'Entries (total)')       = $rechte.Count
    (T 'Ordner' 'Folder')                  = (T 'Hauptkalender' 'Main calendar')
}
$basis = BerichtSchreiben -Name (T 'Kalender-Check' 'Calendar check') -Datei 'Kalender-Check' -Bezug 'Exchange' -BezugName (T 'Geprüft' 'Checked') -Umgebung $umgebung `
    -Abschnitte ([ordered]@{ (T 'Alle Kalenderrechte' 'All calendar permissions') = $liste }) `
    -Fuss (T 'Grundlage sind die Berechtigungen des Hauptkalenders jedes Postfachs. Der Bericht ersetzt kein Berechtigungskonzept und erhebt keinen Anspruch auf Vollständigkeit. Er enthält Namen und Adressen - bitte vertraulich behandeln. Die Vollversion (Kalenderrechte-Bericht) liest zusätzlich "Gelöschte Elemente", löst Gruppen auf, erkennt Mehrfach-Einträge, schlägt Gruppen vor und erzeugt ein prüfbares Änderungsskript.' 'The basis is the permissions of the main calendar of each mailbox. The report does not replace a permission concept and makes no claim to completeness. It contains names and addresses - please treat as confidential. The full version (calendar permissions report) additionally reads "Deleted Items", resolves groups, detects duplicate entries, suggests groups, and generates a verifiable change script.')
if ($Csv) { SpaltenAnzeige $liste | Export-Csv -Path "${basis}_Kalenderrechte.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8 }
if ($PassThru) { $Befunde }
