# ============================================================
#  ODCA Document Log Webpart - One-Click Setup Script
#  Run as Administrator in PowerShell
# ============================================================

$ErrorActionPreference = "Stop"
$REPO_URL   = "https://github.com/mai-sc/document-log-webpart.git"
$CLONE_DIR  = "$env:USERPROFILE\Desktop\document-log-webpart"
$NODE_VER   = "18"

function Write-Step($msg) { Write-Host "`n>> $msg" -ForegroundColor Cyan }
function Write-Ok($msg)   { Write-Host "   $msg"   -ForegroundColor Green }
function Write-Skip($msg) { Write-Host "   $msg"   -ForegroundColor Yellow }

# ── 1. Install Git ──────────────────────────────────────────
Write-Step "Checking Git..."
if (Get-Command git -ErrorAction SilentlyContinue) {
    Write-Ok "Git already installed: $(git --version)"
} else {
    Write-Step "Installing Git via winget..."
    winget install --id Git.Git -e --accept-source-agreements --accept-package-agreements
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path", "User")
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        Write-Host "   Git installed but not in PATH yet. Please restart PowerShell and re-run this script." -ForegroundColor Red
        exit 1
    }
    Write-Ok "Git installed."
}

# ── 2. Install nvm-windows ─────────────────────────────────
Write-Step "Checking nvm..."
if (Get-Command nvm -ErrorAction SilentlyContinue) {
    Write-Ok "nvm already installed."
} else {
    Write-Step "Installing nvm-windows via winget..."
    winget install --id CoreyButler.NVMforWindows -e --accept-source-agreements --accept-package-agreements
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path", "User")
    if (-not (Get-Command nvm -ErrorAction SilentlyContinue)) {
        Write-Host "   nvm installed but not in PATH yet. Please restart PowerShell and re-run this script." -ForegroundColor Red
        exit 1
    }
    Write-Ok "nvm installed."
}

# ── 3. Install & use Node 18 ───────────────────────────────
Write-Step "Setting up Node $NODE_VER..."
$installedVersions = nvm list 2>&1
if ($installedVersions -match $NODE_VER) {
    Write-Skip "Node $NODE_VER already installed."
} else {
    nvm install $NODE_VER
    Write-Ok "Node $NODE_VER installed."
}
nvm use $NODE_VER
$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
            [System.Environment]::GetEnvironmentVariable("Path", "User")
Write-Ok "Using Node $(node -v)"

# ── 4. Install global tools ────────────────────────────────
Write-Step "Installing gulp-cli globally..."
if (Get-Command gulp -ErrorAction SilentlyContinue) {
    Write-Skip "gulp-cli already installed."
} else {
    npm install -g gulp-cli
    Write-Ok "gulp-cli installed."
}

Write-Step "Trusting dev certificate..."
gulp trust-dev-cert

# ── 5. Clone or pull repo ──────────────────────────────────
if (Test-Path "$CLONE_DIR\.git") {
    Write-Step "Repo already cloned. Pulling latest from main..."
    Set-Location $CLONE_DIR
    git checkout main
    git pull origin main
    Write-Ok "Up to date."
} else {
    Write-Step "Cloning repo to Desktop..."
    git clone $REPO_URL $CLONE_DIR
    Set-Location $CLONE_DIR
    Write-Ok "Cloned."
}

# ── 6. Install dependencies ────────────────────────────────
Write-Step "Installing npm dependencies..."
npm install
Write-Ok "Dependencies installed."

# ── 7. Serve ────────────────────────────────────────────────
Write-Step "Starting gulp serve..."
Write-Host "   The workbench will open in your browser shortly.`n" -ForegroundColor Magenta
gulp serve
