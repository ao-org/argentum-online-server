## Por favor considera apoyarnos en https://www.patreon.com/nolandstudios

Como correr mi server:
Crear una nueva carpeta llamada AO20 y entrar a la misma, en ella seguir lo siguientes pasos:

1- Clonar `argentum20-server`

2- Renombrar carpeta `argentum20-server` a `re20-server`.

3- Renombrar el archivo `Example.Server.ini` a `Server.ini`

4- Clonar `Recursos`

5- Abrir Visual Basic 6 como administrador

6- Abrir el archivo `Server.VBP`

# Pull Requests

Before make a `git commit` please run the file git_ignore_case.sh to avoid false changes in the PR.

# Server AO20
Código fuente del servidor de Argentum20

![image](https://i.ibb.co/gFDn3SG/AO20-drawio-2.png)

# Requisitos

## Database
- http://www.ch-werner.de/sqliteodbc/
- https://dev.mysql.com/get/Downloads/Connector-ODBC/8.0/mysql-connector-odbc-8.0.22-win32.msi

## Networking
- Liberia de networking - https://github.com/Wolftein/Aurora.Network

Registrar manualmente libreria Aurora.Network.dll 
Abrir CMD como Administrador `regsvr32 Aurora.Network.dll`

## Cryptography
CryptoSys is used in AO20 to cipher sensitive data.

- https://www.cryptosys.net/api.html 

Please note this is not free software and you will have to buy your own license to use CryptoSys

## Microsoft Visual C++ Redistributable
- https://docs.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist?view=msvc-170

# Logs
Los logs estan en la carpeta de Logs, Errores y en Windows Events.



