# AI 办公助手 v2.0 - 后端服务启动脚本
Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  AI 办公助手 v2.0 - 后端服务启动中..." -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

$serverDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $serverDir

# 检查 node_modules
if (-not (Test-Path "node_modules")) {
    Write-Host "[1/2] 正在安装依赖..." -ForegroundColor Yellow
    npm install
    if ($LASTEXITCODE -ne 0) {
        Write-Host "安装失败，请检查网络连接后重试" -ForegroundColor Red
        Read-Host "按 Enter 退出"
        exit 1
    }
    Write-Host "依赖安装完成" -ForegroundColor Green
}

# 检查 .env
if (-not (Test-Path ".env")) {
    Write-Host "[提示] 未发现 .env 文件，正在从模板创建..." -ForegroundColor Yellow
    Copy-Item ".env.example" ".env"
    Write-Host "请编辑 .env 文件填入你的 API Key 后重新启动" -ForegroundColor Yellow
    Invoke-Item ".env"
    Read-Host "按 Enter 退出"
    exit 1
}

Write-Host "[2/2] 启动服务..." -ForegroundColor Yellow
Write-Host ""
node src/server.js

Read-Host "按 Enter 退出"
