@echo off
chcp 65001 >nul
title AI 办公助手 - 后端服务

cd /d "%~dp0"

echo ═══════════════════════════════════════
echo  AI 办公助手 v2.0 - 后端服务启动中...
echo ═══════════════════════════════════════
echo.

:: 检查 node_modules
if not exist "node_modules" (
    echo [1/2] 正在安装依赖...
    call npm install
    if %errorlevel% neq 0 (
        echo 安装失败，请检查网络连接后重试
        pause
        exit /b 1
    )
    echo 依赖安装完成
)

:: 检查 .env 文件
if not exist ".env" (
    echo [提示] 未发现 .env 文件，正在从模板创建...
    copy .env.example .env >nul
    echo 请编辑 .env 文件填入你的 API Key 后重新启动
    notepad .env
    pause
    exit /b 1
)

echo [2/2] 启动服务...
echo.
node src/server.js

pause
