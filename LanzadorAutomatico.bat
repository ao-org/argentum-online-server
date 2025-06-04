@echo off
setlocal

REM Carpeta base fija
set "BASE=C:\AO20"

REM Ejecutar el servidor
start "" "%BASE%\argentum-online-server\Server.exe"

REM Ejecutar el creador de índices
start "" "%BASE%\argentum-online-creador-indices\CrearIndices.bat"

REM Esperar 45 segundos antes de iniciar el cliente
timeout /t 45 /nobreak > nul

REM Ejecutar el cliente
start "" "%BASE%\argentum-online-client\argentum.exe"

REM Cerrar la ventana de este BAT
exit
