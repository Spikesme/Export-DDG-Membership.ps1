<#
Export-DDG-Membership-Interactive.ps1

Exportiert die Mitgliedschaft ausgewählter Dynamic Distribution Groups (DDGs)
interaktiv in zwei CSVs:
- Summary je Gruppe
- Detailzeilen je Mitglied

Funktionen:
- Fragt interaktiv nach einem Gruppennamen/-muster
- Zeigt passende DDGs an
- Auswahl einzelner, mehrerer oder aller Gruppen
- Optional: nur UserMailbox
- Optional: Company-Filter

Kompatibel mit Windows PowerShell 5.1 + ExchangeOnlineManagement.
#>

[CmdletBinding()]
param()

# ------------------------------------------------------------
# Hilfsfunktionen
# ------------------------------------------------------------
function Write-Info {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Cyan
}

function Write-Ok {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Green
}

function Write-WarnMsg {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Yellow
}

function Read-YesNo {
    param(
        [string]$Prompt,
        [bool]$Default = $false
    )

    $suffix = if ($Default) { " [J/n]" } else { " [j/N]" }

    while ($true) {
        $inputValue = Read-Host ($Prompt + $suffix)

        if ([string]::IsNullOrWhiteSpace($inputValue)) {
            return $Default
        }

        switch ($inputValue.Trim().ToLower()) {
            'j'   { return $true }
            'ja'  { return $true }
            'y'   { return $true }
            'yes' { return $true }
            'n'   { return $false }
            'nein'{ return $false }
            'no'  { return $false }
            default {
                Write-WarnMsg "Bitte J oder N eingeben."
            }
        }
    }
}

function Read-GroupSelection {
    param(
        [array]$Groups
    )

    while ($true) {
        Write-Host ""
        Write-Info "Gefundene Gruppen:"
        for ($i = 0; $i -lt $Groups.Count; $i++) {
            Write-Host ("[{0}] {1}" -f ($i + 1), $Groups[$i].Name)
        }

        Write-Host ""
        Write-Host "Auswahlmöglichkeiten:"
        Write-Host "  - Eine Nummer:         z.B. 1"
        Write-Host "  - Mehrere Nummern:     wähle mehrere Nummern. z.B. 1,4,7"
        Write-Host "  - Bereich:             2-5"
        Write-Host "  - Alles:               a"
        Write-Host ""

        $selection = Read-Host "Welche Gruppe(n) möchtest du exportieren?"

        if ([string]::IsNullOrWhiteSpace($selection)) {
            Write-WarnMsg "Bitte eine Auswahl eingeben."
            continue
        }

        $selection = $selection.Trim()

        if ($selection -match '^(a|all|alle)$') {
            return ,$Groups
        }

        $indices = New-Object System.Collections.Generic.List[int]
        $valid = $true

        $parts = $selection -split '\s*,\s*'
        foreach ($part in $parts) {
            if ($part -match '^\d+$') {
                $index = [int]$part
                if ($index -lt 1 -or $index -gt $Groups.Count) {
                    $valid = $false
                    break
                }
                if (-not $indices.Contains($index - 1)) {
                    $indices.Add($index - 1)
                }
            }
            elseif ($part -match '^(\d+)\s*-\s*(\d+)$') {
                $start = [int]$Matches[1]
                $end   = [int]$Matches[2]

                if ($start -gt $end) {
                    $tmp = $start
                    $start = $end
                    $end = $tmp
                }

                if ($start -lt 1 -or $end -gt $Groups.Count) {
                    $valid = $false
                    break
                }

                for ($i = $start; $i -le $end; $i++) {
                    if (-not $indices.Contains($i - 1)) {
                        $indices.Add($i - 1)
                    }
                }
            }
            else {
                $valid = $false
                break
            }
        }

        if (-not $valid -or $indices.Count -eq 0) {
            Write-WarnMsg "Ungültige Auswahl. Beispiel: 1,3,5 oder 2-4 oder a"
            continue
        }

        $selectedGroups = foreach ($idx in ($indices | Sort-Object)) {
            $Groups[$idx]
        }

        return ,$selectedGroups
    }
}

# ------------------------------------------------------------
# EXO verbinden (robust)
# ------------------------------------------------------------
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
} catch {
    throw "ExchangeOnlineManagement konnte nicht geladen werden: $($_.Exception.Message)"
}

$connected = $false
try {
    $ci = Get-ConnectionInformation -ErrorAction Stop
    if ($ci) { $connected = $true }
} catch {
    $connected = $false
}

if (-not $connected) {
    Write-Info "Keine aktive EXO-Verbindung gefunden. Verbinde zu Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false
}

# ------------------------------------------------------------
# Interaktive Eingaben
# ------------------------------------------------------------
Write-Host ""
Write-Host "=========================================" -ForegroundColor Magenta
Write-Host "   Export DDG Membership - Interaktiv    " -ForegroundColor Magenta
Write-Host "=========================================" -ForegroundColor Magenta
Write-Host ""

$nameLike = Read-Host "Suchbegriff für Gruppenname (z. B. Abteilung* oder *Einkauf*)"

Write-Info "Suche Gruppen..."
$groups = Get-DynamicDistributionGroup -ResultSize Unlimited |
          Where-Object { $_.Name -like $nameLike } |
          Sort-Object Name

if (-not $groups -or $groups.Count -eq 0) {
    Write-WarnMsg "Keine Gruppen gefunden für Muster: $nameLike"
    return
}

Write-Ok ("Gefundene Gruppen: {0}" -f $groups.Count)

$selectedGroups = Read-GroupSelection -Groups $groups

Write-Host ""
Write-Info ("Ausgewählte Gruppen: {0}" -f $selectedGroups.Count)
$selectedGroups | ForEach-Object { Write-Host (" - {0}" -f $_.Name) }

$onlyUserMailbox = Read-YesNo -Prompt "Nur UserMailbox exportieren?" -Default $false

$onlyCompany = Read-Host "Optional Company-Filter (leer lassen = kein Filter)"
if (-not [string]::IsNullOrWhiteSpace($onlyCompany)) {
    Write-Info "Company-Filter aktiv: $onlyCompany*"
} else {
    $onlyCompany = $null
}

# ------------------------------------------------------------
# Exportpfade
# ------------------------------------------------------------
$now = Get-Date -Format 'yyyyMMdd_HHmmss'

if ($selectedGroups.Count -eq 1) {
    $safeGroupName = ($selectedGroups[0].Name -replace '[\\/:*?"<>| ]','_')
    $summaryPath = ".\DDG_Membership_Summary_{0}_{1}.csv" -f $safeGroupName, $now
    $membersPath = ".\DDG_Members_{0}_{1}.csv" -f $safeGroupName, $now
    $errorsPath  = ".\DDG_Membership_Errors_{0}_{1}.log" -f $safeGroupName, $now
} else {
    $summaryPath = ".\DDG_Membership_Summary_{0}.csv" -f $now
    $membersPath = ".\DDG_Members_{0}.csv" -f $now
    $errorsPath  = ".\DDG_Membership_Errors_{0}.log" -f $now
}

$sumRows = New-Object System.Collections.Generic.List[object]
$memRows = New-Object System.Collections.Generic.List[object]

# ------------------------------------------------------------
# Verarbeitung
# ------------------------------------------------------------
foreach ($g in $selectedGroups) {
    try {
        $filter = [string]$g.RecipientFilter

        if ([string]::IsNullOrWhiteSpace($filter)) {
            $msg = "Leerer RecipientFilter – übersprungen: $($g.Name)"
            Write-WarnMsg $msg
            Add-Content -Path $errorsPath -Value ("[{0}] {1}" -f (Get-Date), $msg)
            continue
        }

        Write-Info ("Verarbeite: {0}" -f $g.Name)

        # Mitglieder gemäß DDG-Filter
        $members = Get-Recipient -Filter $filter -ResultSize Unlimited -ErrorAction Stop

        # Optional: Export-seitige Zusatzfilter
        if ($onlyUserMailbox) {
            $members = $members | Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' }
        }

        if ($onlyCompany) {
            $members = $members | Where-Object { $_.Company -like ($onlyCompany + '*') }
        }

        $type = if ($g.Name -match '\.Directs$') {
            'Directs'
        }
        elseif ($g.Name -match '\.Alle$') {
            'Alle'
        }
        else {
            'Other'
        }

        $root = ($g.Name -replace '^CLD\.DDG\.','' -replace '\..*$','')
        $memberCount = ($members | Measure-Object).Count

        $sumRows.Add([pscustomobject]@{
            GroupName        = $g.Name
            GroupPrimarySmtp = $g.PrimarySmtpAddress
            GroupType        = $type
            Root             = $root
            FilterLength     = $filter.Length
            MemberCount      = $memberCount
        })

        foreach ($m in $members) {
            $memRows.Add([pscustomobject]@{
                GroupName            = $g.Name
                GroupType            = $type
                Root                 = $root
                MemberName           = $m.DisplayName
                MemberPrimarySmtp    = $m.PrimarySmtpAddress
                RecipientTypeDetails = $m.RecipientTypeDetails
                Company              = $m.Company
                Department           = $m.Department
                Title                = $m.Title
            })
        }

        Write-Ok ("{0}: {1} Mitglied(er)" -f $g.Name, $memberCount)
    }
    catch {
        $msg = "[{0}] {1}: {2}" -f (Get-Date), $g.Name, $_.Exception.Message
        Write-WarnMsg $msg
        Add-Content -Path $errorsPath -Value $msg
    }
}

# ------------------------------------------------------------
# CSV schreiben
# ------------------------------------------------------------
$sumRows | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
$memRows | Export-Csv -Path $membersPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Ok ("Summary: {0}" -f (Resolve-Path $summaryPath))
Write-Ok ("Members: {0}" -f (Resolve-Path $membersPath))

if (Test-Path $errorsPath) {
    Write-WarnMsg ("Errors : {0}" -f (Resolve-Path $errorsPath))
}