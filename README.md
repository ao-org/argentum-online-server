## ⚔️ Por favor considera apoyarnos en [https://www.patreon.com/nolandstudios](https://www.patreon.com/nolandstudios) ⚔️

![ao20 logo](https://www.ao20.com.ar/_nuxt/img/ao20_logo_sm.d4333ec.png)

# 🛡️ Server AO20
Código fuente del servidor de Argentum20

# 🛡️ Cómo correr mi server:
Crear una nueva carpeta llamada `C:\AO20` y entrar a la misma, en ella seguir lo siguientes pasos:

1. Clonar repositorio `git clone https://github.com/ao-org/argentum20-server.git`

2. Renombrar carpeta `argentum20-server` a `re20-server`.

3. Renombrar el archivo `Example.Server.ini` a `Server.ini`

4. Clonar `Recursos`

5. Abrir Visual Basic 6 como administrador

6. Abrir el archivo `Server.VBP`

# 🛡️ Pull Requests

Before make a `git commit` please run the file `git_ignore_case.sh` to avoid false changes in the PR.

# 🛡️ Requisitos

## Database
- [http://www.ch-werner.de/sqliteodbc/](http://www.ch-werner.de/sqliteodbc/)
- [https://dev.mysql.com/get/Downloads/Connector-ODBC/8.0/mysql-connector-odbc-8.0.22-win32.msi](https://dev.mysql.com/get/Downloads/Connector-ODBC/8.0/mysql-connector-odbc-8.0.22-win32.msi)

## Networking
- Liberia de networking - [https://github.com/Wolftein/Aurora.Network](https://github.com/Wolftein/Aurora.Network)

Registrar manualmente libreria Aurora.Network.dll 
Abrir CMD como Administrador `regsvr32 Aurora.Network.dll`

## Cryptography
CryptoSys is used in AO20 to cipher sensitive data.

- [https://www.cryptosys.net/api.html](https://www.cryptosys.net/api.html)

Please note this is not free software and you will have to buy your own license to use CryptoSys

## Microsoft Visual C++ Redistributable
- [https://docs.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist?view=msvc-170](https://docs.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist?view=msvc-170)

# 🛡️ Logs
Los logs están en la carpeta de Logs, Errores y en Windows Events.

