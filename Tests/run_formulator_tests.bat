@echo off
REM Build and run the Formulator tests. Uses the VS 2022 Developer Command
REM Prompt environment if invoked there; otherwise tries vcvars64.bat.
REM
REM Links Formulator.cpp against the stub CGrid32Mgr members at the bottom of
REM FormulatorTests.cpp rather than Grid32Mgr.cpp, so the parser is testable
REM without standing up a window, GDI+ and the edit subclass.

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
cl.exe /nologo /EHsc /std:c++17 /W3 /Zi /MDd /DUNICODE /D_UNICODE ^
    /Fobuild\ /Fdbuild\ /Febuild\FormulatorTests.exe ^
    FormulatorTests.cpp ..\Grid32\Formulator.cpp
if errorlevel 1 (
    echo Build failed.
    exit /b 1
)

echo.
build\FormulatorTests.exe
exit /b %errorlevel%
