@echo off
REM Build and run the Grid32 helper tests. Uses the Developer Command
REM Prompt environment if invoked from one; otherwise locates Visual Studio
REM with vswhere and runs vcvars64.bat itself.

setlocal
set TESTS_DIR=%~dp0
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

if not exist build mkdir build
cl.exe /nologo /EHsc /std:c++17 /W3 /Zi /MDd /Fobuild\ /Fdbuild\ /Febuild\Grid32Tests.exe Grid32Tests.cpp
if errorlevel 1 (
    echo Build failed.
    exit /b 1
)

echo.
build\Grid32Tests.exe
exit /b %errorlevel%
