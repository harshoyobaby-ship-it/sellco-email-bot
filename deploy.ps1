# Push SELLCO Email Bot to GitHub (run after: gh auth login)

$ErrorActionPreference = "Stop"

$gh = Get-Command gh -ErrorAction SilentlyContinue
if (-not $gh) {
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path", "User")
}

Write-Host "`n=== SELLCO Email Bot — GitHub deploy ===`n" -ForegroundColor Cyan

gh auth status 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "You are not signed in to GitHub yet." -ForegroundColor Yellow
    Write-Host "Run this first, then run this script again:`n"
    Write-Host "  gh auth login`n" -ForegroundColor Green
    Write-Host "Choose: GitHub.com -> HTTPS -> Login with browser`n"
    exit 1
}

$repoName = "sellco-email-bot"
$existing = git remote get-url origin 2>$null

if (-not $existing) {
    Write-Host "Creating GitHub repo: $repoName ..." -ForegroundColor Cyan
    gh repo create $repoName --public --source=. --remote=origin --push --description "SELLCO bulk email bot with Excel upload"
} else {
    Write-Host "Remote already set: $existing" -ForegroundColor Cyan
    Write-Host "Pushing to GitHub ..."
    git push -u origin main
}

if ($LASTEXITCODE -ne 0) {
    Write-Host "`nPush failed. Check errors above." -ForegroundColor Red
    exit 1
}

$repoUrl = gh repo view --json url -q .url 2>$null
Write-Host "`nDone! Code is on GitHub." -ForegroundColor Green
if ($repoUrl) { Write-Host "Repo: $repoUrl`n" }

Write-Host "=== Next: deploy the web app (free) ===`n" -ForegroundColor Cyan
Write-Host "1. Open https://share.streamlit.io"
Write-Host "2. Sign in with GitHub"
Write-Host "3. New app -> pick repo '$repoName' -> main file: streamlit_app.py"
Write-Host "4. App settings -> Secrets -> paste from .streamlit\secrets.toml.example"
Write-Host "5. Deploy and share the link with your team`n"

Start-Process "https://share.streamlit.io"
