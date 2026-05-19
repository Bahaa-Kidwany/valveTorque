@echo off
echo ======================================================================
echo   Creating Clean Isolated Environment for Fugro JDF Analyzer
echo ======================================================================
echo.

set PY_EXE="C:\Users\Kidwany\anaconda3\python.exe"

if not exist venv (
    echo [1/3] Creating a clean virtual environment in 'venv' folder...
    %PY_EXE% -m venv venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
) else (
    echo [1/3] Clean virtual environment 'venv' already exists.
)

echo.
echo [2/3] Activating environment and installing ONLY target dependencies...
call venv\Scripts\activate.bat

:: Upgrade pip first
python -m pip install --upgrade pip

:: Install strictly the required lightweight packages
python -m pip install PyQt5 openpyxl pandas numpy matplotlib pymupdf pyinstaller

echo.
echo [3/3] Compiling standalone executable in clean environment...
echo This will be extremely fast and produce a very small executable size!
echo.

pyinstaller --noconsole --onefile --clean --icon="example-assets\app_icon.ico" --add-data "example-assets\extracted_img_p1_0.png;." --add-data "example-assets\app_icon.ico;." --name="FugroJDFAnalyzer" app.py

echo.
echo Deactivating virtual environment...
call deactivate

echo.
echo ======================================================================
echo   Build Complete! Standalone executable is in 'dist\FugroJDFAnalyzer.exe'
echo   Executable size is now beautifully optimized and lightweight!
echo ======================================================================
pause
