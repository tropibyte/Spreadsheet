@echo off
REM Build and run the Grid32 live-window tests against a built Grid32.dll.
REM   run_live_tests.bat [path\to\Grid32.dll] [outdir]
REM The DLL defaults to the x64 Release build of Grid32.vcxproj, which is what
REM CI builds; outdir, if given, receives the rendered images as BMPs.
REM Uses the Developer Command Prompt environment if invoked from one;
REM otherwise locates Visual Studio with vswhere and runs vcvars64.bat itself.

setlocal
set TESTS_DIR=%~dp0
set "GRID32_DLL=%~f1"
if "%~1"=="" set "GRID32_DLL=%TESTS_DIR%..\Grid32\x64\Release\Grid32.dll"
set "OUT_DIR=%~f2"
cd /d "%TESTS_DIR%"

REM Already in a Developer Command Prompt (or a CI job that ran vcvars)?
where cl.exe >nul 2>&1
if not errorlevel 1 goto :have_compiler

REM Otherwise ask vswhere where Visual Studio is, rather than guessing at
REM hard-coded edition paths -- works for Community/Professional/Enterprise,
REM any year, and on CI runners.
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

if not exist "%GRID32_DLL%" (
    echo Error: "%GRID32_DLL%" not found. Build Grid32 first, or pass the DLL's path.
    exit /b 1
)

if not exist build mkdir build
cl.exe /nologo /EHsc /std:c++17 /W4 /MD /Fobuild\ /Febuild\Grid32LiveTests.exe Grid32LiveTests.cpp user32.lib gdi32.lib
if errorlevel 1 (
    echo Build failed.
    exit /b 1
)

echo.
build\Grid32LiveTests.exe "%GRID32_DLL%" "%OUT_DIR%"
exit /b %errorlevel%
