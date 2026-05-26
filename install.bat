@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   飞书 Bot 一键安装
echo ========================================
echo.

echo [1/4] 安装 Python 依赖...
pip install -r "%~dp0requirements.txt" -q
if %errorlevel% neq 0 (
    echo 错误: pip 安装失败，请确认已安装 Python 并勾选了 Add to PATH
    pause
    exit /b 1
)
echo 依赖安装完成
echo.

echo [2/4] 配置 .env 凭证...
if not exist "%~dp0.env" (
    echo FEISHU_APP_ID=cli_ > "%~dp0.env"
    echo FEISHU_APP_SECRET= >> "%~dp0.env"
)
start notepad "%~dp0.env"
echo 请在弹出的记事本中填入飞书 App ID 和 Secret，保存后关闭，然后按任意键继续...
pause >nul
echo.

echo [3/4] 注册开机自启...
powershell -Command "$p='%~dp0'; Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'FeishuBot' -Value ($p+'start_silent.vbs')"
if %errorlevel% neq 0 (
    echo 开机自启注册失败，请右键手动运行 install_autostart.ps1
) else (
    echo 开机自启已注册
)
echo.

echo [4/4] 适配启动脚本为系统 Python...
echo @echo off > "%~dp0start_auto.bat"
echo chcp 65001 ^>nul >> "%~dp0start_auto.bat"
echo cd /d "%%~dp0" >> "%~dp0start_auto.bat"
echo python main.py >> "%~dp0start_auto.bat"

powershell -Command "(Get-Content '%~dp0start.ps1') -replace '\.\\\.venv\\Scripts\\python\.exe', 'python' | Set-Content '%~dp0start.ps1' -Encoding UTF8"
echo 启动脚本已适配
echo.

echo ========================================
echo   安装完成！
echo   双击 start.vbs 即可手动启动
echo   重启电脑后会自动静默启动
echo ========================================
pause
