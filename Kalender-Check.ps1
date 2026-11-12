#Requires -Version 5.1
<#
.SYNOPSIS
    Kalender-Check: zeigt für Exchange Server (on-prem), wer auf welchen Kalender zugreifen darf, und meldet einige
    Auffälligkeiten (Standard-Recht zu weit, Anonym-Zugriff, Schreibrechte). Ergebnis ist ein HTML-Bericht.

.DESCRIPTION
    Das Skript arbeitet ausschließlich LESEND. Es verwendet nur Get-Befehle von Exchange (Get-Mailbox,
    Get-MailboxFolderStatistics, Get-MailboxFolderPermission), ändert keine Berechtigung und überträgt nichts.

    Gelesen wird der Hauptkalender jedes Postfachs. Der Ordner wird über den Ordnertyp gesucht, deshalb funktioniert das
    Skript in jeder Postfachsprache; es liest nacheinander und ist bei sehr vielen Postfächern entsprechend langsam
    (grob eine Sekunde je Postfach). Am einfachsten direkt auf einem Exchange-Server starten: Das Skript lädt die
    Exchange-Verwaltungsshell selbst.

    Voraussetzungen: Windows PowerShell 5.1, Exchange Server 2016/2019/SE, zum Lesen genügt die Rolle
    "View-Only Recipients" (z. B. über "View-Only Organization Management").

.PARAMETER Ausgabe
    Pfad der HTML-Datei. Standard: Kalender-Check_Exchange_<Datum>.html im aktuellen Ordner.

.PARAMETER Postfach
    Nur diese Postfächer prüfen (Adressen). Standard: alle Benutzer-, Gemeinsamen, Raum- und Gerätepostfächer.

.PARAMETER ExchangeServer
    Name oder FQDN eines Exchange-Servers für eine Remote-Sitzung (Kerberos), wenn auf dem Rechner keine Exchange-
    Verwaltungstools installiert sind.

.PARAMETER Csv
    Schreibt zusätzlich je Prüfpunkt mit Treffern eine CSV-Datei und die Liste aller Kalenderrechte neben den Bericht.

.PARAMETER NichtOeffnen
    Öffnet den fertigen Bericht nicht automatisch im Browser.

.PARAMETER PassThru
    Gibt die Befunde zusätzlich als Objekte zurück (für eigene Auswertungen).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Kalender-Check.ps1

.EXAMPLE
    .\Kalender-Check.ps1 -ExchangeServer ex01.firma.local -Csv -Ausgabe C:\Berichte\kalender.html

.NOTES
    IT Admin Tools by MM - Kalender-Check Version 1.0 (11.2026), Lizenz: MIT (siehe LICENSE)
    Kein Ersatz für ein Berechtigungskonzept. Nutzung auf eigene Verantwortung.
#>
[CmdletBinding()]
param(
    [string]$Ausgabe,
    [string[]]$Postfach,
    [string]$ExchangeServer,
    [switch]$Csv,
    [switch]$NichtOeffnen,
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'
$Version = '1.0'

# Kommalisten: Beim Start über "powershell -File" kommt -Postfach a,b als ein Text an. Deshalb an Kommas trennen.
$nurAuswahl = $PSBoundParameters.ContainsKey('Postfach')
if ($Postfach) { $Postfach = @($Postfach | ForEach-Object { $_ -split '\s*,\s*' } | Where-Object { $_ }) }
if ($nurAuswahl -and -not $Postfach) { throw 'Bei -Postfach wurde kein Postfach angegeben. Ohne den Parameter werden alle Postfächer geprüft.' }

# ------------------------------------------------------------------ Hilfsfunktionen
function Datum($d, $leer = 'nie') { if ($d) { ([datetime]$d).ToString('dd.MM.yyyy') } else { $leer } }
function JaNein($b) { if ($b) { 'Ja' } else { 'Nein' } }
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
        $sicher | Microsoft.PowerShell.Utility\Export-Csv -Path $Path -NoTypeInformation:$NoTypeInformation -Delimiter $Delimiter -Encoding $Encoding
    }
}

# Tabellen: ConvertTo-Html und Export-Csv nehmen die Spalten des ersten Objekts. Haben die Zeilen einer Liste verschiedene
# Eigenschaften (z. B. Fehlzeile neben normalen Treffern), gingen Spalten verloren: alle Namen der Reihe nach vereinigen.
function Einheitlich($liste) {
    $liste = @($liste | Where-Object { $null -ne $_ })
    if ($liste.Count -lt 2) { return $liste }
    $namen = New-Object System.Collections.Generic.List[string]
    foreach ($o in $liste) { foreach ($p in $o.PSObject.Properties.Name) { if (-not $namen.Contains($p)) { $namen.Add($p) } } }
    if ($namen.Count -eq @($liste[0].PSObject.Properties.Name).Count) { return $liste }
    @($liste | Select-Object -Property $namen)
}
# Sortierschluessel fuer Nummern wie K2, K10: erst Buchstaben, dann die Zahl als Zahl
function IdZahl($id) { $m = [regex]::Match([string]$id, '\d+'); if ($m.Success) { [int]$m.Value } else { 0 } }
function IdText($id) { [regex]::Replace([string]$id, '\d+', '') }

$Befunde = New-Object System.Collections.Generic.List[object]

# Platzhalterzeilen ("nicht geprüft") erkennt das Skript nur an den vollständigen, festen Texten, die es selbst erzeugt - keine Teilstring-Suche
# in Konto-, Objekt- oder Beschreibungstexten, die ein Angreifer beeinflussen könnte
$PlatzhalterTexte = @('Abfrage fehlgeschlagen - nicht geprüft', 'Berechtigungen nicht lesbar - nicht geprüft', 'SYSVOL nicht lesbar - nicht geprüft', 'WMI-Filter nicht lesbar - nicht geprüft', 'Kein Domänencontroller ermittelbar - SYSVOL nicht geprüft', 'Inhalt des Ordners nicht lesbar - Richtlinie nicht geprüft')
function IstPlatzhalter($o) {
    # Ein Skript kann $script:PlatzhalterSpalten setzen: dann zählen nur diese Spalten (Namensspalten sind von Benutzern beeinflussbar)
    foreach ($pr in $o.PSObject.Properties) {
        if ($script:PlatzhalterSpalten -and ($pr.Name -notin $script:PlatzhalterSpalten)) { continue }
        if (($pr.Value -is [string]) -and ($PlatzhalterTexte -ccontains $pr.Value)) { return $true }
    }
    ("$($o.Feststellung)" -clike 'Nicht geprüft - *')
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
function NichtGeprueft([string]$IdMuster, [string[]]$Namen, [string]$Verweis, [string]$Spalte = 'Server', [string]$Grund = 'Abfrage nicht möglich') {
    foreach ($x in ($Befunde | Where-Object { $_.Id -match $IdMuster })) {
        $war = $x.Anzahl
        $zeilen = @($Namen | ForEach-Object { [pscustomobject][ordered]@{ $Spalte = $_; Feststellung = "Nicht geprüft - $Grund (siehe $Verweis)"; Details = '' } })
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
    param([string]$Name, [string]$Datei, [string]$Bezug, [string]$BezugName = 'Domäne', [System.Collections.IDictionary]$Umgebung,
        [System.Collections.IDictionary]$Abschnitte = @{}, [System.Collections.IDictionary]$Anhang = @{}, [string]$Fuss)
    Write-Host '  Bericht schreiben ...'
    $jetzt = Get-Date
    $ziel = $script:Ausgabe
    if (-not $ziel) { $ziel = Join-Path (Get-Location).Path ("{0}_{1}_{2:yyyy-MM-dd_HHmm}.html" -f $Datei, ($Bezug -replace '[^\w.\-]+', '-'), $jetzt) }
    $ziel = [System.IO.Path]::GetFullPath($ziel)
    $ordner = Split-Path $ziel -Parent
    if (-not (Test-Path $ordner)) { New-Item -ItemType Directory -Path $ordner -Force | Out-Null }

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
    [void]$h.Append("<!DOCTYPE html><html lang='de'><head><meta charset='utf-8'><title>$(Html $Name) $(Html $Bezug)</title><style>$css</style></head><body><div class='seite'>")
    [void]$h.Append("<h1>$(Html $Name)</h1><div class='meta'>$(Html $BezugName) <b>$(Html $Bezug)</b> · erstellt am $($jetzt.ToString('dd.MM.yyyy HH:mm')) von $(Html "$env:USERDOMAIN\$env:USERNAME") auf $(Html $env:COMPUTERNAME)</div>")
    [void]$h.Append("<div class='kacheln'>")
    foreach ($s in 'Hoch', 'Mittel', 'Niedrig', 'Info', 'OK') {
        $text = @{ Hoch = 'Befunde mit hoher Priorität'; Mittel = 'mit mittlerer Priorität'; Niedrig = 'mit niedriger Priorität'; Info = 'Hinweise und nicht geprüft'; OK = 'Prüfpunkte ohne Befund' }[$s]
        [void]$h.Append("<div class='kachel $s'><b>$($zahl[$s])</b><span>$text</span></div>")
    }
    [void]$h.Append("</div><h2>Umgebung</h2><table><tr><th>Eigenschaft</th><th>Wert</th></tr>")
    foreach ($k in $Umgebung.Keys) { [void]$h.Append("<tr><td>$(Html $k)</td><td>$(Html $Umgebung[$k])</td></tr>") }
    [void]$h.Append("</table>")

    [void]$h.Append("<h2>Übersicht der Prüfpunkte</h2><table><tr><th>Nr.</th><th>Bereich</th><th>Prüfpunkt</th><th>Treffer</th><th>Bewertung</th></tr>")
    foreach ($x in $Befunde) {
        [void]$h.Append("<tr><td>$($x.Id)</td><td>$(Html $x.Bereich)</td><td><a href='#$($x.Id)'>$(Html $x.Titel)</a></td><td>$($x.Anzahl)</td><td><span class='marke m-$($x.Stufe)'>$($x.Stufe)</span></td></tr>")
    }
    [void]$h.Append("</table>")

    foreach ($k in $Abschnitte.Keys) {
        [void]$h.Append("<h2>$(Html $k)</h2>")
        $inhalt = $Abschnitte[$k]
        if ($inhalt -is [string]) { [void]$h.Append($inhalt) }
        elseif (@($inhalt).Count) { [void]$h.Append(((Einheitlich $inhalt | ConvertTo-Html -Fragment) -join "`n")) }
        else { [void]$h.Append("<p class='klein'>Keine Einträge.</p>") }
    }

    [void]$h.Append("<h2>Befunde</h2>")
    $offen = @($Befunde | Where-Object { $_.Stufe -ne 'OK' } | Sort-Object { $rang[$_.Stufe] }, { IdText $_.Id }, { IdZahl $_.Id })
    foreach ($x in $offen) {
        [void]$h.Append("<div class='befund $($x.Stufe)' id='$($x.Id)'><h3><span class='marke m-$($x.Stufe)'>$($x.Stufe)</span> $($x.Id) · $(Html $x.Titel) <span class='klein'>($($x.Anzahl) Treffer)</span></h3>")
        [void]$h.Append("<p>$(Html $x.Erklaerung)</p><p class='emp'><b>Empfehlung:</b> $(Html $x.Empfehlung)</p>")
        [void]$h.Append(((Einheitlich $x.Objekte | ConvertTo-Html -Fragment) -join "`n"))
        [void]$h.Append("</div>")
    }
    if (-not $offen.Count) { [void]$h.Append("<p>Keine Auffälligkeiten gefunden.</p>") }
    $ok = @($Befunde | Where-Object Stufe -eq 'OK')
    if ($ok.Count) {
        [void]$h.Append("<h2>Ohne Befund</h2><table><tr><th>Nr.</th><th>Prüfpunkt</th></tr>")
        foreach ($x in $ok) { [void]$h.Append("<tr id='$($x.Id)'><td>$($x.Id)</td><td>$(Html $x.Titel)</td></tr>") }
        [void]$h.Append("</table>")
    }
    foreach ($k in $Anhang.Keys) { [void]$h.Append("<h2>$(Html $k)</h2><pre>$(Html $Anhang[$k])</pre>") }
    [void]$h.Append("<div class='fuss'>$(Html $Name) Version $script:Version · IT Admin Tools by MM · Das Skript hat ausschließlich gelesen und nichts verändert. $(Html $Fuss)</div>")
    [void]$h.Append("</div></body></html>")
    [System.IO.File]::WriteAllText($ziel, $h.ToString(), (New-Object System.Text.UTF8Encoding($true)))

    $basis = [System.IO.Path]::ChangeExtension($ziel, $null).TrimEnd('.')
    if ($script:Csv) {
        foreach ($x in ($Befunde | Where-Object { $_.Anzahl -gt 0 })) {
            Einheitlich $x.Objekte | Export-Csv -Path "${basis}_$($x.Id).csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8
        }
    }
    Write-Host ''
    Write-Host ("Fertig: {0} hoch, {1} mittel, {2} niedrig, {3} ohne Befund." -f $zahl.Hoch, $zahl.Mittel, $zahl.Niedrig, $zahl.OK) -ForegroundColor Green
    Write-Host "Bericht: $ziel"
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
        Write-Host 'Lade Exchange-Verwaltungsshell ...' -ForegroundColor Cyan
        $ErrorActionPreference = 'Continue'
        try { . $rex *> $null; Connect-ExchangeServer -auto -ClientApplication:ManagementShell *> $null }
        catch { Write-Warning "Die Exchange-Verwaltungsshell ließ sich nicht laden: $($_.Exception.Message)" }
        $ErrorActionPreference = 'Stop'
    }
}
if (-not (ExchangeBereit)) {
    if (-not $ExchangeServer) { throw 'Die Exchange-Befehle sind nicht verfügbar. Bitte auf einem Exchange-Server starten oder -ExchangeServer angeben.' }
    Write-Host "Baue Remote-Sitzung zu $ExchangeServer auf ..." -ForegroundColor Cyan
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$ExchangeServer/PowerShell/" -Authentication Kerberos -SessionOption (New-PSSessionOption -OpenTimeout 60000)
    Import-PSSession -Session $session -DisableNameChecking -AllowClobber -CommandName Get-Mailbox, Get-MailboxFolderStatistics, Get-MailboxFolderPermission 3>$null | Out-Null
}

# ------------------------------------------------------------------ Lesen
Write-Host 'Kalender-Check - es wird nur gelesen, nichts verändert.' -ForegroundColor Cyan
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
Write-Host "$($mbx.Count) Postfächer werden gelesen ..." -ForegroundColor Cyan

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
        if (-not $f) { throw 'Kein Standardkalender gefunden (Postfach evtl. nie angemeldet).' }
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
    catch { $nichtGeprueft.Add([pscustomobject][ordered]@{ Postfach = $smtp; Feststellung = 'Nicht geprüft - Kalenderrechte nicht lesbar'; Details = Kurz $_.Exception.Message 160 }) }
}
Write-Progress -Activity 'Lese Kalenderrechte' -Completed
if ($session) { Remove-PSSession $session }

# ------------------------------------------------------------------ Prüfungen
$schreib = 'Author', 'PublishingAuthor', 'NonEditingAuthor', 'Contributor', 'Editor', 'PublishingEditor', 'Owner', 'FolderOwner'
function IstSchreibrecht([string]$r) { foreach ($x in ($r -split ',\s*')) { if ($x -in $schreib -or $x -match '^(CreateItems|EditAllItems|EditOwnedItems|DeleteAllItems|DeleteOwnedItems)$') { return $true } }; $false }
function PostfachTyp([string]$t) { switch ($t) { 'UserMailbox' { 'Benutzer' } 'SharedMailbox' { 'Gemeinsam' } 'RoomMailbox' { 'Raum' } 'EquipmentMailbox' { 'Gerät' } default { $t } } }
function Zeile($r, $zusatz = @{}) {
    $o = [ordered]@{ Postfach = $r.Postfach; Adresse = $r.Adresse; 'Postfach-Typ' = (PostfachTyp $r.Typ); Berechtigter = $r.Berechtigter; Recht = $r.Recht }
    foreach ($k in $zusatz.Keys) { $o[$k] = $zusatz[$k] }
    [pscustomobject]$o
}
$nurFrei = 'None', 'AvailabilityOnly'

Befund 'K1' 'Standardzugriff' 'Alle Benutzer dürfen mehr als Frei/Gebucht sehen (Standard-Recht)' 'Mittel' `
    @($rechte | Where-Object { $_.Art -eq 'Default' -and $_.Recht -notin $nurFrei } | Sort-Object Postfach | ForEach-Object { Zeile $_ }) `
    'Das Standard-Recht gilt für jeden Benutzer der Organisation ohne eigenen Eintrag. Ab "LimitedDetails" sehen alle Betreff und Ort, ab "Reviewer" den ganzen Termin. Bei Kalendern von Geschäftsführung, Personal oder Betriebsrat ist das selten gewollt.' `
    'Standard-Recht auf "AvailabilityOnly" (nur Frei/Gebucht) setzen, wo mehr nicht nötig ist, und gezielt Berechtigte eintragen.'
Befund 'K2' 'Standardzugriff' 'Anonymer Zugriff auf einen Kalender' 'Hoch' `
    @($rechte | Where-Object { $_.Art -eq 'Anonymous' -and $_.Recht -ne 'None' } | Sort-Object Postfach | ForEach-Object { Zeile $_ }) `
    'Anonymer Zugriff erlaubt Personen ohne Anmeldung, den Kalender zu sehen, zum Beispiel über eine veröffentlichte Adresse im Internet.' `
    'Anonym auf "None" setzen, wenn nicht ausdrücklich eine Veröffentlichung gewollt ist.'
Befund 'K3' 'Schreibrechte' 'Personen oder Gruppen mit Schreibrecht auf einen fremden Kalender' 'Info' `
    @($rechte | Where-Object { $_.Art -in 'Intern', 'External' -and (IstSchreibrecht $_.Recht) } | Sort-Object Postfach, Berechtigter | ForEach-Object { Zeile $_ @{ Art = $(if ($_.Gruppe) { 'Gruppe' } else { 'Person' }) } }) `
    'Mit Schreibrecht können Termine angelegt, geändert und (je nach Recht) gelöscht werden. Das ist bei Stellvertretern und Sekretariaten üblich, sollte aber bekannt und gewollt sein.' `
    'Liste mit den Verantwortlichen durchgehen und nicht mehr benötigte Rechte entfernen.'
Befund 'K4' 'Aufräumen' 'Einträge für Konten, die nicht mehr auflösbar sind' 'Niedrig' `
    @($rechte | Where-Object { $_.Art -eq 'Unknown' } | Sort-Object Postfach | ForEach-Object { Zeile $_ }) `
    'Das Konto wurde gelöscht oder ist nicht erreichbar. Der Eintrag bleibt als Karteileiche stehen und verwirrt bei der Fehlersuche.' `
    'Eintrag nach einem Blick entfernen (Remove-MailboxFolderPermission), das Skript ändert nichts.'
Befund 'K5' 'Abfrage' 'Postfächer, deren Kalender nicht gelesen werden konnte' 'Info' $nichtGeprueft.ToArray() `
    'Für diese Postfächer fehlt die Auswertung; sie sind in den übrigen Prüfpunkten nicht enthalten.' `
    'Meldung lesen (Rolle, Postfach nie angemeldet?) und das Skript erneut starten.'

# ------------------------------------------------------------------ Bericht
$liste = $rechte | Sort-Object Postfach, Berechtigter | ForEach-Object {
    $istGruppe = $_.Gruppe   # innerhalb von switch ist $_ der geprüfte Wert, deshalb vorher merken
    [pscustomobject][ordered]@{ Postfach = $_.Postfach; Adresse = $_.Adresse; 'Postfach-Typ' = (PostfachTyp $_.Typ); Berechtigter = $_.Berechtigter; Art = $(switch ($_.Art) { 'Default' { 'Standard (alle)' } 'Anonymous' { 'Anonym' } 'Unknown' { 'unbekannt' } 'External' { 'extern' } default { if ($istGruppe) { 'Gruppe' } else { 'Person' } } }); Recht = $_.Recht }
}
$umgebung = [ordered]@{
    'Postfächer geprüft'      = "$([math]::Max(0, $mbx.Count - ($nichtGeprueft.Count - $nichtGefunden.Count))) von $($mbx.Count + $nichtGefunden.Count)"
    'Einträge (gesamt)'       = $rechte.Count
    'Ordner'                  = 'Hauptkalender'
}
$basis = BerichtSchreiben -Name 'Kalender-Check' -Datei 'Kalender-Check' -Bezug 'Exchange' -BezugName 'Geprüft' -Umgebung $umgebung `
    -Abschnitte ([ordered]@{ 'Alle Kalenderrechte' = $liste }) `
    -Fuss 'Grundlage sind die Berechtigungen des Hauptkalenders jedes Postfachs. Der Bericht ersetzt kein Berechtigungskonzept und erhebt keinen Anspruch auf Vollständigkeit. Er enthält Namen und Adressen - bitte vertraulich behandeln. Die Vollversion (Kalenderrechte-Bericht) liest zusätzlich "Gelöschte Elemente", löst Gruppen auf, erkennt Mehrfach-Einträge, schlägt Gruppen vor und erzeugt ein prüfbares Änderungsskript.'
if ($Csv) { $liste | Export-Csv -Path "${basis}_Kalenderrechte.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8 }
if ($PassThru) { $Befunde }
