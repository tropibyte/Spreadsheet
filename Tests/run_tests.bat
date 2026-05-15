@echo off
REM Build and run Grid32 tests. Uses the VS 2022 Developer Command Prompt
REM environment if invoked there; otherwise tries vcvars64.bat.

setlocal
set TESTS_DIR=%~dp0
cd /d "%TESTS_DIR%"

where cl.exe >nul 2>&1
if errorlevel 1 (
    if exist "%ProgramFiles%\Microsoft Visual Studio\18\Enterprise\VC\Auxiliary\Build\vcvars64.bat" (
        call "%ProgramFiles%\Microsoft Visual Studio\18\Enterprise\VC\Auxiliary\Build\vcvars64.bat" >nul
    ) else if exist "%ProgramFiles%\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" (
        call "%ProgramFiles%\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" >nul
    ) else (
        echo Error: cl.exe not in PATH and vcvars64.bat not found in known locations.
        exit /b 1
    )
)

if not exist build mkdir build
cl.exe /nologo /EHsc /std:c++17 /W3 /Zi /MDd /Fobuild\ /Fdbuild\ /Febuild\Grid32Tests.exe Grid32Tests.cpp
if errorlevel 1 (
    echo Build failed.
    exit /b 1
)

echo.
build\Grid32Tests.exe
exit /b %errorlevel%
