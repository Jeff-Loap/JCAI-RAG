@echo off
setlocal

cd /d "%~dp0"

set "PYTHON_EXE=D:\vng\python.exe"
if exist "%PYTHON_EXE%" goto run_app

set "PYTHON_EXE=python"

:run_app
echo Starting Legal RAG Agent...
echo Working directory: %cd%
echo Python: %PYTHON_EXE%
echo.

"%PYTHON_EXE%" -m streamlit run app.py
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo.
    echo Failed to start the app. Exit code: %EXIT_CODE%
    echo Check whether Streamlit and project dependencies are installed.
    pause
)

endlocal
