@echo off
echo ========================================
echo 舱单文件处理系统 - Electron 打包脚本
echo ========================================
echo.

REM 检查 Node.js 是否安装
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [错误] 未找到 Node.js，请先安装 Node.js
    pause
    exit /b 1
)

echo [步骤 1/3] 安装依赖...
call pnpm install

echo.
echo [步骤 2/3] 打包 Electron 应用...
call npx electron-builder --win --x64

echo.
echo [步骤 3/3] 完成!
echo.
echo 打包后的文件位于 dist 目录中
echo - 安装程序: dist\舱单文件处理系统 Setup 1.0.0.exe
echo - 便携版: dist\舱单文件处理系统 1.0.0.exe
echo.
pause
