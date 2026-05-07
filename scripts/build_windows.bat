@echo off
setlocal

cd /d "%~dp0\.."

python -m pip install --upgrade pip
python -m pip install -e .[build]
pyinstaller --clean --noconfirm kaoqin.spec

echo.
echo Build finished. Output: dist\kaoqin\kaoqin.exe
