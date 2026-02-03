@echo off
echo ==========================================
echo   Trial Reset - 打包工具
echo ==========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python！
    echo 请从 https://python.org 安装 Python
    pause
    exit /b 1
)

echo [1/3] 安装依赖...
pip install -r requirements.txt --quiet

echo [2/3] 打包可执行文件...
pyinstaller --onefile --windowed --name="TrialReset" --icon=assets/icon.ico --add-data "assets;assets" --clean --noconfirm main.py

echo [3/3] 完成！
echo.
echo ==========================================
echo   输出: dist\TrialReset.exe
echo ==========================================
pause
