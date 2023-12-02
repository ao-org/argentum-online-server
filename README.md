### ⚔️ Por favor considera apoyarnos en [https://www.patreon.com/nolandstudios](https://www.patreon.com/nolandstudios) ⚔️ 

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

<a href="https://imgbb.com/"><img src="https://i.ibb.co/QfZznrw/Screenshot-2023-12-02-211157.png" alt="Precommit-hook" border="0"></a>

We have a pre-commit hook for the project, Visual Basic 6 IDE it changes the names of the variables and it makes the Pull Requests very difficult to understand.

Please run the following commands with `git bash` or the client you are using.

```
chmod +x .githooks/pre-commit
git config core.hooksPath .githooks
```

Basically the pre-commit hook runs when you make a `git commit` and it will run the file `git_ignore_case.sh` to avoid false changes in the Pull Request. Is not perfect but it helps a lot. Please send the Pull Requests with only the neccesary code to be reviewed.

In case you have problems setting locally your pre-commit hook you can run the file `git_ignore_case.sh` by just doing double click.

# 🛡️ Requisitos

## Database
- [http://www.ch-werner.de/sqliteodbc/](http://www.ch-werner.de/sqliteodbc/)
Mejorar velocidad de la base de datos con `PRAGMA journal_mode=WAL;`
Para mas detalles visiten https://www.sqlite.org/wal.html

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
