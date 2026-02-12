<#
Export-DDG-Membership.ps1
Exportiert die Mitgliedschaft ALLER vorhandenen CLD.DDG.F1*-Gruppen (Dynamic Distribution Groups)
in zwei CSVs: Summary je Gruppe + Detailzeilen je Mitglied.
Angabe der DDG kann im Skript gemacht werden!


Standard: exportiert genau die Ergebnisse des Gruppenfilters (Get-Recipient -Filter <RecipientFilter>).
Optional: -OnlyUserMailbox und/oder -OnlyCompany 'DEW21' filtern die Export-Zeilen zusätzlich.

Kompatibel mit Windows PowerShell 5.1 + ExchangeOnlineManagement.
#>

[CmdletBinding()]
param(
  [string]$NameLike = 'CLD.DDG.F1',
  [switch]$OnlyUserMailbox,
  [string]$OnlyCompany
)

# -- EXO verbinden (robust) --
try { Import-Module ExchangeOnlineManagement -ErrorAction Stop } catch {}
$connected = $false
try {
  $ci = Get-ConnectionInformation -ErrorAction Stop
  if($ci){ $connected = $true }
} catch { $connected = $false }
if (-not $connected) { Connect-ExchangeOnline -ShowBanner:$false }

$now = Get-Date -Format 'yyyyMMdd_HHmmss'
$summaryPath = ".\DDG_Membership_Summary_$now.csv"
$membersPath = ".\DDG_Members_$now.csv"
$errorsPath  = ".\DDG_Membership_Errors_$now.log"

$sumRows = New-Object System.Collections.Generic.List[object]
$memRows = New-Object System.Collections.Generic.List[object]

# Alle CLD.DDG.*.*-Gruppen einsammeln
$groups = Get-DynamicDistributionGroup -ResultSize Unlimited |
          Where-Object { $_.Name -like $NameLike } |
          Sort-Object Name

Write-Host ("Gefundene Gruppen: {0}" -f $groups.Count)

foreach($g in $groups){
  try {
    $filter = [string]$g.RecipientFilter
    if ([string]::IsNullOrWhiteSpace($filter)) {
      $msg = "Leerer RecipientFilter – übersprungen: $($g.Name)"
      Write-Warning $msg
      Add-Content -Path $errorsPath -Value ("[{0}] {1}" -f (Get-Date), $msg)
      continue
    }

    # Mitglieder gemäß Filter ermitteln
    $members = Get-Recipient -Filter $filter -ResultSize Unlimited -ErrorAction Stop

    # Optional: Export-Filter anwenden (nur für CSV-Ausgabe – Gruppenfilter bleibt unberührt)
    if ($OnlyUserMailbox) {
      $members = $members | Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' }
    }
    if ($OnlyCompany) {
      # begins-with für flexible Abgleiche (Export-seitig ok)
      $members = $members | Where-Object { $_.Company -like ($OnlyCompany + '*') }
    }

    $type = if ($g.Name -match '\.Directs$') { 'Directs' } elseif ($g.Name -match '\.Alle$') { 'Alle' } else { 'Other' }
    $root = ($g.Name -replace '^CLD\.DDG\.','' -replace '\..*$','')

    # Summary-Zeile
    $sumRows.Add([pscustomobject]@{
      GroupName        = $g.Name
      GroupPrimarySmtp = $g.PrimarySmtpAddress
      GroupType        = $type
      Root             = $root
      FilterLength     = ([string]$filter).Length
      MemberCount      = ($members | Measure-Object).Count
    })

    # Mitglieder-Zeilen
    foreach($m in $members){
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

    Write-Host ("{0}: {1}" -f $g.Name, ($members | Measure-Object).Count)

  } catch {
    $msg = "[{0}] {1}: {2}" -f (Get-Date), $g.Name, $_.Exception.Message
    Write-Warning $msg
    Add-Content -Path $errorsPath -Value $msg
  }
}

# CSV schreiben
$sumRows | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
$memRows | Export-Csv -Path $membersPath -NoTypeInformation -Encoding UTF8

Write-Host ("Summary: {0}" -f (Resolve-Path $summaryPath)) -ForegroundColor Green
Write-Host ("Members: {0}" -f (Resolve-Path $membersPath)) -ForegroundColor Green
if (Test-Path $errorsPath) {
  Write-Host ("Errors : {0}" -f (Resolve-Path $errorsPath)) -ForegroundColor Yellow
}
