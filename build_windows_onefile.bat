@echo off
echo ===================================================
echo   HopePharma Windows Builder
echo ===================================================
echo.
echo This script will build the Windows executable.
echo Please ensure Python is installed and added to PATH.
echo.

pip install pyinstaller reportlab pillow

echo.
echo Building HopePharma...
echo.

pyinstaller HopePharma.spec --noconfirm

echo.
echo ===================================================
echo   BUILD COMPLETE
echo ===================================================
echo.
echo You can find the new app in the "dist" folder.
echo.
pause
