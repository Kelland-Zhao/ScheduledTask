param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Message
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

git add -A
git status --short

if (-not (git diff --cached --name-only)) {
    Write-Host "Nothing to commit" -ForegroundColor Yellow
    exit 0
}

git commit -m $Message
git push origin HEAD

Write-Host "Pushed: $Message" -ForegroundColor Green
