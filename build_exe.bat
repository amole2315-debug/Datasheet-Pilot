@echo off
setlocal

set "CONDA_BAT=conda"
if exist "%USERPROFILE%\anaconda3\condabin\conda.bat" set "CONDA_BAT=%USERPROFILE%\anaconda3\condabin\conda.bat"
if exist "%USERPROFILE%\miniconda3\condabin\conda.bat" set "CONDA_BAT=%USERPROFILE%\miniconda3\condabin\conda.bat"

call "%CONDA_BAT%" activate MPX2
if errorlevel 1 (
    echo Failed to activate conda environment: MPX2
    exit /b 1
)

set "TCL_LIBRARY=%CONDA_PREFIX%\Library\lib\tcl8.6"
set "TK_LIBRARY=%CONDA_PREFIX%\Library\lib\tk8.6"

python -m PyInstaller DatasheetCorrection.spec --clean --noconfirm
if errorlevel 1 (
    echo Build failed.
    exit /b 1
)

echo.
echo Done: dist\DatasheetCorrection.exe
endlocal
