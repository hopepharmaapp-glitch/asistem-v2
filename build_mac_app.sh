#!/bin/bash
echo "Building HopePharma for macOS..."
# Ensure dependencies are installed
python3 -m pip install pyinstaller reportlab pillow
# Run PyInstaller
pyinstaller HopePharma.spec --noconfirm
echo "Build complete. Check the dist folder."
