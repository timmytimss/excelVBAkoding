# Publiserer en ny EMI-versjon fra det offentlige prosjektrepoet timmytimss/excelVBAkoding.
# EMI (v2.9.0.0+) leser "Excel Macro Installer/versjon.json" fra main ved oppstart og tilbyr
# «Oppdater EMI» nar pakkeversjonen der er hoyere enn den installerte. Oppdateringen kopierer
# KUN exe, Macros\*.ps1, Resources\, Data\macro-metadata.json og versjon.json (hviteliste i EMI).
#
# Forutsetning: kodeendringene (makroer/exe) er allerede kopiert inn i git-klonen og committet.
# Skriptet oker bare versjon.json, committer den og pusher - det er selve «publiseringen».
#
# Bruk:  powershell -ExecutionPolicy Bypass -File Publiser-EMI.ps1 -Endringer "Kort tekst om hva som er nytt"
# Pakkeversjon = aaaa.MM.dd.N (N teller opp ved flere publiseringer samme dag).
# Laget 2026-10-08 (kontoflytting fase 7). Se Excel VBA Koding.html (EMI -> Oppdatering fra GitHub).
param(
    [Parameter(Mandatory = $true)][string]$Endringer,
    [string]$Repo = (Join-Path $env:USERPROFILE "Documents\excelVBAkoding-git"),
    [string]$Versjon
)
$ErrorActionPreference = 'Stop'
function Invoke-RepoGit { & git.exe -C $Repo @args; if ($LASTEXITCODE -ne 0) { throw "git $($args -join ' ') feilet ($LASTEXITCODE)" } }

$emi = Join-Path $Repo "Excel Macro Installer"
$kildeVersjon = ([regex]::Match((Get-Content -Raw -Encoding UTF8 (Join-Path $emi "Installer\Excel Macro Installer.ps1")), '\$LauncherVersion\s*=\s*"([^"]+)"')).Groups[1].Value
$exeVersjon = [System.Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $emi "Excel Macro Installer.exe")).FileVersion
if ($kildeVersjon -ne $exeVersjon) { throw "Exe-en ($exeVersjon) matcher ikke `$LauncherVersion i kilden ($kildeVersjon). Rekompiler og commit exe-en forst." }

Invoke-RepoGit pull --ff-only origin main
$fil = Join-Path $emi "versjon.json"
$forrige = if (Test-Path $fil) { (Get-Content -Raw -Encoding UTF8 $fil | ConvertFrom-Json).versjon } else { $null }
if (-not $Versjon) {
    $idag = Get-Date -Format "yyyy.MM.dd"; $n = 1
    if ($forrige -and $forrige.StartsWith($idag + ".")) { $n = [int]($forrige.Split('.')[3]) + 1 }
    $Versjon = "$idag.$n"
}
if ($forrige -and ([version]$Versjon -le [version]$forrige)) { throw "Ny versjon $Versjon er ikke hoyere enn forrige ($forrige)." }

$json = [ordered]@{ versjon = $Versjon; dato = (Get-Date -Format "yyyy-MM-dd"); emi = $kildeVersjon; endringer = $Endringer } | ConvertTo-Json
[System.IO.File]::WriteAllText($fil, $json, (New-Object System.Text.UTF8Encoding($false)))
# Speil til OneDrive-kilden hvis skriptet kjores derfra (holder kilde og repo like).
$speil = Join-Path $PSScriptRoot "..\..\Excel Macro Installer\versjon.json"
if ((Test-Path (Split-Path $speil -Parent)) -and ((Resolve-Path (Split-Path $speil -Parent)).Path -ne (Resolve-Path $emi).Path)) {
    [System.IO.File]::WriteAllText($speil, $json, (New-Object System.Text.UTF8Encoding($false)))
}
Invoke-RepoGit add -- "Excel Macro Installer/versjon.json"
Invoke-RepoGit commit -q -m "EMI-pakke $Versjon (EMI $kildeVersjon): $Endringer"
Invoke-RepoGit push -q origin main
Write-Host "Publisert EMI-pakke $Versjon (EMI $kildeVersjon). Installasjoner tilbyr oppdatering ved neste oppstart (GitHub kan bruke opptil ca. 5 min)."