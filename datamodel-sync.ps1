# =============================================================================
# datamodel-sync.ps1
# Dataverse Solution Export & Unpack - Manual Sync Script
#
# Usage:
#   .\datamodel-sync.ps1                          # sync obě solution
#   .\datamodel-sync.ps1 -Target Base             # pouze TSANETBaseSolution
#   .\datamodel-sync.ps1 -Target Dynamics         # pouze TSANETDynamicsSolution
#   .\datamodel-sync.ps1 -SkipCommit              # sync bez git commitu
# =============================================================================

param(
    [ValidateSet("All", "Base", "Dynamics")]
    [string]$Target = "All",

    [switch]$SkipCommit
)

$ErrorActionPreference = "Stop"

# =============================================================================
# KONFIGURACE - upravte dle prostředí
# =============================================================================

$Config = @{
    Base = @{
        Environment  = "https://org65ec56f5.crm.dynamics.com"   # TODO: doplň URL prostředí 1
        SolutionName = "TSANET"
        OutputFolder = "DataModel/TSANETBaseSolution"
        ZipName      = "TSANETBaseSolution.zip"
    }
    Dynamics = @{
        Environment  = "https://orgce975757.crm.dynamics.com"
        SolutionName = "TSANETDynamicsCS"
        OutputFolder = "DataModel/TSANETDynamicsSolution"
        ZipName      = "TSANETDynamicsSolution.zip"
    }
}

# =============================================================================
# POMOCNÉ FUNKCE
# =============================================================================

function Write-Step {
    param([string]$Message)
    Write-Host "`n=== $Message ===" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host "✓ $Message" -ForegroundColor Green
}

function Write-Fail {
    param([string]$Message)
    Write-Host "✗ $Message" -ForegroundColor Red
}

function Test-PacCli {
    try {
        $null = pac --version
        return $true
    } catch {
        return $false
    }
}

function Sync-Solution {
    param(
        [string]$Environment,
        [string]$SolutionName,
        [string]$OutputFolder,
        [string]$ZipName
    )

    $ZipPath = Join-Path $RepoRoot "exports/$ZipName"
    $FolderPath = Join-Path $RepoRoot $OutputFolder

    Write-Step "Exporting '$SolutionName' from $Environment"

    # Export ze Dataverse
    $ExportArgs = @(
        "--name", $SolutionName,
        "--path", $ZipPath,
        "--managed", "false",
        "--environment", $Environment
    )
    pac solution export @ExportArgs

    Write-Success "Export dokončen: $ZipName"

    Write-Step "Unpacking '$SolutionName' to $OutputFolder"

    # Unpack do složky
    $UnpackArgs = @(
        "--zipfile", $ZipPath,
        "--folder", $FolderPath,
        "--processCanvasApps", "true",
        "--allowDelete", "true"
    )
    pac solution unpack @UnpackArgs

    Write-Success "Unpack dokončen: $OutputFolder"
}

# =============================================================================
# MAIN
# =============================================================================

# Zjisti kořen repozitáře (skript je v /scripts, repo root je o úroveň výš)
$RepoRoot = Split-Path $PSScriptRoot -Parent

Write-Host "`n=============================================" -ForegroundColor Yellow
Write-Host " Dataverse Solution Sync" -ForegroundColor Yellow
Write-Host " Repo: $RepoRoot" -ForegroundColor Yellow
Write-Host " Target: $Target" -ForegroundColor Yellow
Write-Host "=============================================`n" -ForegroundColor Yellow

# Ověř PAC CLI
Write-Step "Kontrola PAC CLI"
if (-not (Test-PacCli)) {
    Write-Fail "PAC CLI není nainstalováno."
    Write-Host "Instalace: winget install Microsoft.PowerAppsCLI" -ForegroundColor Yellow
    exit 1
}
Write-Success "PAC CLI nalezeno: $(pac --version)"

# Ověř git
Write-Step "Kontrola git stavu"
Set-Location $RepoRoot
$GitStatus = git status --porcelain
if ($GitStatus) {
    Write-Host "Pozor: máš necommitované změny:" -ForegroundColor Yellow
    Write-Host $GitStatus
    $Continue = Read-Host "Pokračovat? (y/n)"
    if ($Continue -ne "y") { exit 0 }
}
Write-Success "Git OK"

# Vytvoř exports složku
New-Item -ItemType Directory -Force -Path "$RepoRoot/exports" | Out-Null

# Spusť sync dle parametru Target
try {
    if ($Target -eq "All" -or $Target -eq "Base") {
        $BaseParams = @{
            Environment  = $Config.Base.Environment
            SolutionName = $Config.Base.SolutionName
            OutputFolder = $Config.Base.OutputFolder
            ZipName      = $Config.Base.ZipName
        }
        Sync-Solution @BaseParams
    }

    if ($Target -eq "All" -or $Target -eq "Dynamics") {
        $DynamicsParams = @{
            Environment  = $Config.Dynamics.Environment
            SolutionName = $Config.Dynamics.SolutionName
            OutputFolder = $Config.Dynamics.OutputFolder
            ZipName      = $Config.Dynamics.ZipName
        }
        Sync-Solution @DynamicsParams
    }
} catch {
    Write-Fail "Chyba během sync: $_"
    exit 1
}

# Git commit
if (-not $SkipCommit) {
    Write-Step "Git commit"

    git add DataModel/

    $Changed = git diff --staged --name-only
    if ($Changed) {
        $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"
        git commit -m "chore: sync Dataverse data model [$Timestamp]"
        Write-Success "Commitováno"

        $Push = Read-Host "`nChceš pushovat do GitHubu? (y/n)"
        if ($Push -eq "y") {
            git push
            Write-Success "Pushováno"
        }
    } else {
        Write-Host "`nŽádné změny v DataModel/ – commit přeskočen." -ForegroundColor Yellow
    }
} else {
    Write-Host "`n-SkipCommit: git kroky přeskočeny." -ForegroundColor Yellow
    Write-Host "Zkontroluj změny: git diff DataModel/" -ForegroundColor Yellow
}

Write-Host "`n=============================================" -ForegroundColor Green
Write-Host " Sync dokončen!" -ForegroundColor Green
Write-Host "=============================================`n" -ForegroundColor Green