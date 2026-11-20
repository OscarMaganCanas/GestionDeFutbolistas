@echo off
cd /d "D:\CLAUDE\gestionFutbolistas"
echo Comprobando pptxgenjs...
call npm list pptxgenjs 2>nul | findstr "pptxgenjs" >nul
if errorlevel 1 (
    echo Instalando pptxgenjs...
    call npm install pptxgenjs
)
echo Generando presentacion...
node generar_presentacion.js
echo.
echo Proceso finalizado. Pulsa cualquier tecla para cerrar.
pause
