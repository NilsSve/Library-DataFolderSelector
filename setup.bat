@echo off
setlocal

REM ===========================================================================
REM  Library-DataFolderSelector - one-time setup for STANDALONE development
REM
REM  DataFolderSelector no longer carries DFAbout, RDCToolsLib and vwin32fh as
REM  nested submodules. They are separate sibling libraries, referenced by
REM  DataFolderSelectorDev25.0.sws as ..\DFAbout, ..\RDCToolsLib and ..\vwin32fh.
REM
REM  This script is only for working on DataFolderSelector on its own - it
REM  clones those three as siblings of this folder, which is where the .sws
REM  expects them. When DataFolderSelector is consumed as a library inside
REM  another workspace, that workspace already provides the three (they sit in
REM  the same flat library set), and this script is a no-op there.
REM
REM  Open DataFolderSelectorDev25.0.sws to build. There are two workspace
REM  files, named for their role: DataFolderSelectorLibrary25.0.sws is the
REM  (empty) consumer entry that applications reference; DataFolderSelectorDev
REM  is the one you open to build the library itself. Likewise
REM  RDCToolsLibLibrary25.0.sws is RDCToolsLib's empty consumer entry - vwin32fh
REM  is supplied here as its own sibling, per the flat model.
REM ===========================================================================

cd /d "%~dp0"

echo.
echo === Library-DataFolderSelector standalone setup ===
echo Folder: %CD%
echo.

where git >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Git was not found on your PATH.
    echo         Install Git ^(or the GitHub Desktop app^), reopen the
    echo         command prompt, and run setup.bat again.
    echo.
    pause
    exit /b 1
)

for %%N in (DFAbout RDCToolsLib vwin32fh) do (
    if exist "..\%%N\.git" (
        echo [%%N] already present - pulling latest...
        git -C "..\%%N" pull --ff-only
    ) else (
        echo [%%N] cloning as a sibling...
        git clone https://github.com/NilsSve/Library-%%N.git "..\%%N"
        if errorlevel 1 (
            echo.
            echo [ERROR] Could not clone Library-%%N.
            echo         Check your connection and that you can reach:
            echo           https://github.com/NilsSve/Library-%%N.git
            echo.
            pause
            exit /b 1
        )
    )
)

echo.
echo === Setup complete ===
echo.
echo DFAbout, RDCToolsLib and vwin32fh are now siblings of this folder.
echo Open DataFolderSelectorDev25.0.sws in the Studio and build.
echo.
pause
exit /b 0
