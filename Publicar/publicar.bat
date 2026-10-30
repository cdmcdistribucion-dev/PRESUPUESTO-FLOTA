@echo off
cd /d "%~dp0"
echo Generando datos desde base_excels...
node actualizar_datos.js
if errorlevel 1 (echo.& echo Hubo un error. No se publico nada.& pause & exit /b 1)
git add -A
git commit -m "Actualizacion de datos %date% %time%"
git push
echo.
echo Listo. En 1-2 minutos el tablero publicado muestra los datos nuevos.
pause
