@echo off
REM Windows batch script to run Robot Framework tests

echo Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found!
    echo.
    echo Please install Python 3.7+ from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

echo Python found!
echo.

REM Check if virtual environment exists
if not exist "venv\" (
    echo Creating virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment
        pause
        exit /b 1
    )
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

REM Install/upgrade dependencies
echo Installing dependencies...
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt --quiet

REM Create reports directory
if not exist "reports\" mkdir reports

REM Run tests from tests directory (required for relative resource paths)
echo.
echo Running Robot Framework tests...
echo.
cd tests
robot --outputdir ..\reports test_*.robot
cd ..

REM Check exit code
if errorlevel 1 (
    echo.
    echo Tests completed with failures. Check reports/report.html for details.
) else (
    echo.
    echo All tests passed! Check reports/report.html for details.
)

echo.
echo Opening test report...
if exist "reports\report.html" (
    start reports\report.html
)

pause


