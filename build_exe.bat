@echo off
echo ======================================================================
echo   Creating Clean Isolated Environment for Fugro JDF Analyzer
echo ======================================================================
echo.

if not exist venv (
    echo [1/3] Creating a clean virtual environment in 'venv' folder...
    python -m venv venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment using 'python'. Trying 'py'...
        py -m venv venv
        if errorlevel 1 (
            echo [ERROR] Failed to create virtual environment. Make sure Python is installed.
            pause
            exit /b 1
        )
    )
) else (
    echo [1/3] Clean virtual environment 'venv' already exists.
)

echo.
echo [2/3] Activating environment and installing/upgrading dependencies...
venv\Scripts\python -m pip install --upgrade pip
venv\Scripts\python -m pip install PyQt5 openpyxl pandas numpy matplotlib pymupdf pyinstaller xlrd

echo.
echo [3/3] Compiling standalone executable in clean environment...
echo This will produce a beautifully optimized standalone executable!
echo.

venv\Scripts\pyinstaller --noconsole --onefile --clean --icon="example-assets\app_icon.ico" --add-data "example-assets\extracted_img_p1_0.png;." --add-data "example-assets\app_icon.ico;." --name="FugroJDFAnalyzer" app.py

echo.
echo ======================================================================
echo   Build Complete! Standalone executable is in 'dist\FugroJDFAnalyzer.exe'
echo ======================================================================
