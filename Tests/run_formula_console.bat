@echo off
REM Build and start the formula console: an interactive shell over the formula
REM engine (Tests\FormulaConsole.cpp). Arguments are passed through:
REM
REM   run_formula_console.bat                          interactive
REM   run_formula_console.bat console\examples.txt     run a script; exit code 1
REM                                                    if one of its expects fails
REM   run_formula_console.bat -i my_case.txt           run a script, then stay
REM
REM Script paths are relative to the directory you run this from. Uses the
REM Developer Command Prompt environment if invoked from one; otherwise locates
REM Visual Studio with vswhere, as run_formulator_tests.bat does. Like that
REM suite it links Formulator.cpp against stub CGrid32Mgr members, so no window,
REM GDI+ or DLL is involved.

setlocal
set "TESTS_DIR=%~dp0"
set "CALLER_DIR=%CD%"

where cl.exe >nul 2>&1
if not errorlevel 1 goto :have_compiler

set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" set "VSWHERE=%ProgramFiles%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" (
    echo Error: cl.exe not in PATH and vswhere.exe not found.
    echo Run this from a Developer Command Prompt, or install Visual Studio.
    exit /b 1
)

set "VSPATH="
for /f "usebackq tokens=*" %%i in (`"%VSWHERE%" -latest -products * ^
    -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 ^
    -property installationPath`) do set "VSPATH=%%i"

if not defined VSPATH (
    echo Error: no Visual Studio install with the C++ toolset was found.
    exit /b 1
)
if not exist "%VSPATH%\VC\Auxiliary\Build\vcvars64.bat" (
    echo Error: vcvars64.bat missing under "%VSPATH%".
    exit /b 1
)
call "%VSPATH%\VC\Auxiliary\Build\vcvars64.bat" >nul
:have_compiler

REM Build in the Tests directory (vcvars may have moved us), then go back.
cd /d "%TESTS_DIR%"
if not exist build mkdir build
cl.exe /nologo /EHsc /std:c++17 /W3 /Zi /MDd /DUNICODE /D_UNICODE ^
    /Fobuild\ /Fdbuild\ /Febuild\FormulaConsole.exe ^
    FormulaConsole.cpp ..\Grid32\Formulator.cpp
if errorlevel 1 (
    echo Build failed.
    exit /b 1
)

cd /d "%CALLER_DIR%"
"%TESTS_DIR%build\FormulaConsole.exe" %*
exit /b %errorlevel%
