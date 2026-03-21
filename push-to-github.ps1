# GitHub Push Script
# GitHubリポジトリ作成後に実行してください

Write-Host "Excel MCP Server - GitHub Push Script" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

# Change to project directory
Set-Location "C:\copilot-cli\excel-mcp-server"

Write-Host "Current directory: $(Get-Location)" -ForegroundColor Yellow
Write-Host ""

# Check if remote already exists
$remoteExists = git remote get-url origin 2>$null

if ($remoteExists) {
    Write-Host "Remote 'origin' already exists: $remoteExists" -ForegroundColor Yellow
    $response = Read-Host "Do you want to update it? (y/n)"
    if ($response -eq 'y') {
        git remote set-url origin https://github.com/prizmPrograms/excel-mcp-server.git
        Write-Host "Remote URL updated!" -ForegroundColor Green
    }
} else {
    Write-Host "Adding remote 'origin'..." -ForegroundColor Yellow
    git remote add origin https://github.com/prizmPrograms/excel-mcp-server.git
    Write-Host "Remote added!" -ForegroundColor Green
}

Write-Host ""
Write-Host "Renaming branch to 'main'..." -ForegroundColor Yellow
git branch -M main

Write-Host ""
Write-Host "Pushing to GitHub..." -ForegroundColor Yellow
Write-Host "Note: You may need to enter your GitHub credentials or Personal Access Token" -ForegroundColor Cyan
Write-Host ""

git push -u origin main

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "SUCCESS! Code pushed to GitHub! 🎉" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "1. Visit: https://github.com/prizmPrograms/excel-mcp-server" -ForegroundColor White
    Write-Host "2. Add topics: mcp, mcp-server, excel, vba, windows, automation, japanese" -ForegroundColor White
    Write-Host "3. (Optional) Create Release v1.0.0" -ForegroundColor White
} else {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "ERROR: Push failed!" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Common issues:" -ForegroundColor Yellow
    Write-Host "1. GitHub repository not created yet" -ForegroundColor White
    Write-Host "   -> Create it at: https://github.com/new" -ForegroundColor White
    Write-Host "2. Authentication failed" -ForegroundColor White
    Write-Host "   -> You may need a Personal Access Token" -ForegroundColor White
    Write-Host "   -> Create at: https://github.com/settings/tokens" -ForegroundColor White
    Write-Host "3. Repository name mismatch" -ForegroundColor White
    Write-Host "   -> Ensure repo is named 'excel-mcp-server'" -ForegroundColor White
}

Write-Host ""
Read-Host "Press Enter to exit"
