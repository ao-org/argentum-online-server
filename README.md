### ⚔️ Please consider supporting us at [https://www.patreon.com/nolandstudios](https://www.patreon.com/nolandstudios) ⚔️ 

![ao20 logo](https://www.ao20.com.ar/_nuxt/img/argentum20_logo.562a0aa.png)

# 🛡️ Argentum Online Server
Source code for Argentum Online

# 🛡️ How to Run My Server:
Create a new folder named `C:\AO20` and navigate to it. Follow these steps:

1. Clone the repository `git clone https://github.com/ao-org/argentum-online-server.git`

2. Rename the file `Example.Server.ini` to `Server.ini`

3. Rename the file `Example.feature_toggle.ini` to `feature_toggle.ini`

4. Rename the file `Empty_db.db` to `Database.db`

5. Clone `Recursos` (https://github.com/ao-org/Recursos)

6. Open Visual Basic 6 as an administrator

7. Open the file `Server.VBP`

# 🔛 Feature Toggle/Flag (Turn ON/OFF features)

When introducing new functionality to the server, it should include the capability to be disabled. To achieve this, we implement the feature flags design pattern, which is configured within the file `Example.feature_toggle.ini`.

# 🎬 Game Scenarios
In the following folders, you will find configuration files for events. When programming a new type of event, it must have its own configuration file.
https://github.com/ao-org/Recursos/tree/master/Dat/Scenarios

# 🗄️ SQL Migrations

To modify the schema of the database or make alterations to existing tables, it is essential to create a new SQL migration file within the `ScriptsDB` directory. The project is configured to automatically detect and execute the required migration scripts. This process ensures that the database is systematically updated to reflect the latest schema changes without manual intervention. This approach not only maintains database integrity but also streamlines the update process, enabling seamless transitions between different database schema versions.

# 🛡️ Pull Requests

<a href="https://imgbb.com/"><img src="https://i.ibb.co/QfZznrw/Screenshot-2023-12-02-211157.png" alt="Precommit-hook" border="0"></a>

We have a pre-commit hook for the project, Visual Basic 6 IDE changes the names of the variables and it makes the Pull Requests very difficult to understand.

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
CryptoSys is used in Argentum Online to cipher sensitive data.

- [https://www.cryptosys.net/api.html](https://www.cryptosys.net/api.html)

Please note this is not free software and you will have to buy your own license to use CryptoSys

## Microsoft Visual C++ Redistributable
- [https://docs.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist?view=msvc-170](https://docs.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist?view=msvc-170)

# 🛡️ Logs
Los logs están en la carpeta de Logs, Errores y en Windows Events.


## Star History

<a href="https://star-history.com/#ao-org/argentum-online-server&Date">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="https://api.star-history.com/svg?repos=ao-org/argentum-online-server&type=Date&theme=dark" />
    <source media="(prefers-color-scheme: light)" srcset="https://api.star-history.com/svg?repos=ao-org/argentum-online-server&type=Date" />
    <img alt="Star History Chart" src="https://api.star-history.com/svg?repos=ao-org/argentum-online-server&type=Date" />
  </picture>
</a>


