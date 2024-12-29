@echo off
EventCreate /ID 20 /L Application /T Information /SO Argentum20 /D "Evento de prueba con ID 20 para Argentum20"
if %errorlevel% == 0 (
    echo Evento creado correctamente.
) else (
    echo Error al crear el evento. Código de error: %errorlevel%
)
pause