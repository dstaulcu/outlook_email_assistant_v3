@echo off
REM Outlook Email Assistant - Simple Installer Batch File
REM This batch file provides a simple interface to the PowerShell installer

setlocal

REM Set default values
set ENVIRONMENT=Prd
set SILENT=false
set INSTALL_PATH=%LOCALAPPDATA%\OutlookEmailAssistant

REM Parse command line arguments
:parse_args
if "%~1"=="" goto :run_installer
if /i "%~1"=="--dev" (
    set ENVIRONMENT=Dev
    shift
    goto :parse_args
)
if /i "%~1"=="--prod" (
    set ENVIRONMENT=Prd
    shift
    goto :parse_args
)
if /i "%~1"=="--silent" (
    set SILENT=true
    shift
    goto :parse_args
)
if /i "%~1"=="--uninstall" (
    set UNINSTALL=true
    shift
    goto :parse_args
)
if /i "%~1"=="--help" (
    goto :show_help
)
if /i "%~1"=="-h" (
    goto :show_help
)
if /i "%~1"=="/?" (
    goto :show_help
)
REM Unknown parameter, skip it
shift
goto :parse_args

:show_help
echo.
echo Outlook Email Assistant - Simple Installer
echo ==========================================
echo.
echo Usage: install-outlook-assistant.bat [options]
echo.
echo Options:
echo   --dev         Install from development environment
echo   --prod        Install from production environment (default)
echo   --silent      Install without user prompts
echo   --uninstall   Uninstall the add-in
echo   --help, -h, /?  Show this help message
echo.
echo Examples:
echo   install-outlook-assistant.bat
echo   install-outlook-assistant.bat --dev
echo   install-outlook-assistant.bat --prod --silent
echo   install-outlook-assistant.bat --uninstall
echo.
goto :end

:run_installer
echo.
echo Outlook Email Assistant Installer
echo ==================================
echo.

REM Check for PowerShell
powershell -Command "exit 0" >nul 2>&1
if errorlevel 1 (
    echo ERROR: PowerShell is required but not available.
    echo Please ensure PowerShell is installed and accessible.
    pause
    goto :end
)

REM Get the directory where this batch file is located
set SCRIPT_DIR=%~dp0

REM Build PowerShell command
set PS_SCRIPT=%SCRIPT_DIR%outlook_installer.ps1
set PS_ARGS=-Environment %ENVIRONMENT%

if "%SILENT%"=="true" (
    set PS_ARGS=%PS_ARGS% -Silent
)

if "%UNINSTALL%"=="true" (
    set PS_ARGS=%PS_ARGS% -UninstallOnly
)

REM Check if PowerShell script exists
if not exist "%PS_SCRIPT%" (
    echo ERROR: PowerShell installer script not found: %PS_SCRIPT%
    echo Please ensure outlook_installer.ps1 is in the same directory as this batch file.
    pause
    goto :end
)

REM Run PowerShell script with execution policy bypass
echo Running installer...
echo.
powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %PS_ARGS%

REM Check the exit code
if errorlevel 1 (
    echo.
    echo Installation failed. See error messages above.
    if "%SILENT%"=="false" pause
) else (
    echo.
    echo Installation completed successfully.
    if "%SILENT%"=="false" pause
)

:end
endlocal
