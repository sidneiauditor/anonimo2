@echo off
cd /d "%~dp0"
echo.
echo  Anonimo — servidor local (obrigatorio para PDF + OCR)
echo  Abra no browser: http://localhost:8000/Anonimo.html
echo.
echo  Pressione Ctrl+C para encerrar.
echo.
python -m http.server 8000
