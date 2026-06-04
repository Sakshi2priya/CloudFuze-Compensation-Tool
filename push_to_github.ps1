# =============================================================================
# CloudFuze Compensation Tool — one-shot GitHub push
# =============================================================================
# This script initializes a git repository, commits all your code (excluding
# secrets via .gitignore), and pushes to your GitHub repo. You'll get a
# Git Credential Manager browser prompt for GitHub login the first time only.
#
# Usage:
#   1. Right-click this file → Run with PowerShell
#      (Or open PowerShell here and run: .\push_to_github.ps1)
#   2. A browser tab opens for GitHub sign-in. Sign in with your account
#      (Sakshi2priya). The script does the rest.
# =============================================================================

$ErrorActionPreference = "Stop"
$repoUrl = "https://github.com/Sakshi2priya/CloudFuze-Compensation-Tool.git"
$projectDir = $PSScriptRoot

Write-Host ""
Write-Host "==========================================================================" -ForegroundColor Cyan
Write-Host "  CloudFuze Compensation Tool — GitHub push" -ForegroundColor Cyan
Write-Host "==========================================================================" -ForegroundColor Cyan
Write-Host ""

# --- 1. Verify git is installed -------------------------------------------
Write-Host "[1/7] Checking git..." -ForegroundColor Yellow
try {
    $gitVersion = git --version
    Write-Host "      OK — $gitVersion"
} catch {
    Write-Host ""
    Write-Host "  ERROR: git is not installed. Install Git for Windows first:" -ForegroundColor Red
    Write-Host "  https://git-scm.com/download/win" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# --- 2. Move to project directory -----------------------------------------
Write-Host "[2/7] Project directory: $projectDir" -ForegroundColor Yellow
Set-Location $projectDir

# --- 3. Verify .gitignore exists and excludes .env ------------------------
Write-Host "[3/7] Verifying .gitignore protects secrets..." -ForegroundColor Yellow
if (-not (Test-Path ".gitignore")) {
    Write-Host "  ERROR: .gitignore is missing. Aborting to protect secrets." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}
$gitignoreContent = Get-Content ".gitignore" -Raw
if ($gitignoreContent -notmatch '(?m)^\.env\s*$') {
    Write-Host "  ERROR: .gitignore does not exclude .env. Aborting." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Host "      OK — .env is excluded"

# --- 4. Initialize repo (or reuse if already there) -----------------------
Write-Host "[4/7] Initializing repository..." -ForegroundColor Yellow
if (-not (Test-Path ".git")) {
    git init -b main | Out-Null
    Write-Host "      OK — created new repo"
} else {
    Write-Host "      OK — repo already exists, using it"
}

git config user.name "Sakshi Priya" | Out-Null
git config user.email "Sakshi.Priya@cloudfuze.com" | Out-Null

# --- 5. Stage everything and sanity-check .env is NOT staged --------------
Write-Host "[5/7] Staging files..." -ForegroundColor Yellow
git add .
$stagedEnv = git diff --cached --name-only | Select-String -Pattern '^\.env$' -Quiet
if ($stagedEnv) {
    Write-Host "  STOP — .env was about to be committed. Aborting to protect your secrets." -ForegroundColor Red
    git restore --staged .env 2>$null
    Read-Host "Press Enter to exit"
    exit 1
}
$stagedCount = (git diff --cached --name-only | Measure-Object).Count
Write-Host "      OK — staging $stagedCount files (.env safely excluded)"

# --- 6. Commit -------------------------------------------------------------
Write-Host "[6/7] Creating commit..." -ForegroundColor Yellow
$hasCommits = git rev-parse HEAD 2>$null
if (-not $hasCommits) {
    git commit -m "Initial commit: CloudFuze Compensation Tool" | Out-Null
    Write-Host "      OK — initial commit created"
} else {
    $hasChanges = (git status --porcelain).Length -gt 0
    if ($hasChanges) {
        git commit -m "Update from local machine" | Out-Null
        Write-Host "      OK — update commit created"
    } else {
        Write-Host "      OK — nothing new to commit"
    }
}

# --- 7. Configure remote and push ----------------------------------------
Write-Host "[7/7] Pushing to GitHub..." -ForegroundColor Yellow
Write-Host "      Repo: $repoUrl"
git branch -M main | Out-Null

$existingRemote = git remote get-url origin 2>$null
if (-not $existingRemote) {
    git remote add origin $repoUrl
    Write-Host "      Remote 'origin' added"
} elseif ($existingRemote -ne $repoUrl) {
    git remote set-url origin $repoUrl
    Write-Host "      Remote 'origin' updated to $repoUrl"
} else {
    Write-Host "      Remote 'origin' already set"
}

Write-Host ""
Write-Host "  >>> A browser tab will open for GitHub sign-in (first time only)..." -ForegroundColor Cyan
Write-Host ""

# Actually push. Git Credential Manager handles the browser auth.
git push -u origin main

# --- Done ------------------------------------------------------------------
Write-Host ""
Write-Host "==========================================================================" -ForegroundColor Green
Write-Host "  DONE! Open your repo: https://github.com/Sakshi2priya/CloudFuze-Compensation-Tool" -ForegroundColor Green
Write-Host "==========================================================================" -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to close this window"
