### ⚔️ Please consider supporting us at [https://www.patreon.com/nolandstudios](https://www.patreon.com/nolandstudios) ⚔️ 


![ao logo](https://www.argentumonline.com.ar/_nuxt/img/argentum20_logo.562a0aa.png)

# 🛡️ Argentum Online Server
![image](https://github.com/ao-org/argentum-online-server/assets/5874806/d0f29237-6bd3-4a90-a2d6-fc67b34f1c85)

Important: Do not download the code using the "Download as ZIP" button on GitHub, as this can cause issues with file encoding and may corrupt some files. 

To download the code correctly, use a Git client. The command to clone the repository from the command line is:

```bash
git clone https://www.github.com/ao-org/argentum-online-server
```

# 🛡️ How to Run My Server:
Create a new folder named `C:\AO20` and navigate to it. Follow these steps:

1. Clone the repository `git clone https://github.com/ao-org/argentum-online-server.git`

2. Rename the file `Example.Server.ini` to `Server.ini`

3. Rename the file `Example.feature_toggle.ini` to `feature_toggle.ini`

4. Rename the file `Empty_db.db` to `Database.db`

5. Clone `Recursos` (https://github.com/ao-org/Recursos)

6. Copy `Example.EsArbol.ini` to `Recursos\init\EsArbol.ini` and customize it if you need different tree graphics.

7. Open Visual Basic 6 as an administrator

8. Open the file `Server.VBP`

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

# 🗺️ Maps Limitation 512 with VB6 Debugger

There is a limitation when running and debugging the game within Visual Basic 6. Due to VB6's constraints, it cannot load more than 512 maps during debugging. As a result, maps such as Dungeon Dinosaurios (Map 577) will not function properly and it will throw overflow error

# 🛡️ Requirements

## Database SQLite

### Installing SQLite ODBC Driver for 32-bit Systems

To integrate SQLite with ODBC on a 32-bit system, please download the appropriate driver from the following link:
- [SQLite ODBC Driver - 32 bits](http://www.ch-werner.de/sqliteodbc/sqliteodbc.exe)

### Optimizing Database Performance

To enhance the performance of your SQLite database, consider changing the journal mode to Write-Ahead Logging (WAL) by executing the following SQL command:

```sql
PRAGMA journal_mode=WAL;
```

Write-Ahead Logging can significantly improve the write performance and concurrency of your database. For more information on how WAL mode benefits your database operations, please visit the SQLite WAL documentation:
- [SQLite Write-Ahead Logging](https://www.sqlite.org/wal.html)

This mode enables most read operations to proceed without locking and allows updates to occur without interfering with reads, thus increasing the performance and scalability of your application when using SQLite.

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
* **Location:**
    * Logs folder
    * Errors folder
    * Windows Event Viewer

* **Handling "The description for Event ID:0..." error:**
    Run the `RegistrarEvento.bat` script to resolve this error message within the Windows Event Viewer: 

    ```
    The description for Event ID:0 in Source:'Argentum20' cannot be found. 
    Either the component that raises this event is not installed on your local computer 
    or the installation is corrupted. 
    You can install or repair the component on the local computer. 
    If the event originated on another computer, 
    the display information had to be saved with the event. 
    The following information was included with the event: 
    ```


## Repo Activity
![Alt](https://repobeats.axiom.co/api/embed/f0d51db011fb97750321324a10936f4f7bcf2b87.svg "Repobeats analytics image")

## Star History

<a href="https://star-history.com/#ao-org/argentum-online-server&Date">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="https://api.star-history.com/svg?repos=ao-org/argentum-online-server&type=Date&theme=dark" />
    <source media="(prefers-color-scheme: light)" srcset="https://api.star-history.com/svg?repos=ao-org/argentum-online-server&type=Date" />
    <img alt="Star History Chart" src="https://api.star-history.com/svg?repos=ao-org/argentum-online-server&type=Date" />
  </picture>
</a>

## Thank you

A big thank you ❤️ to these amazing people for contributing to this project!

<a href="https://github.com/ao-org/argentum-online-server/graphs/contributors">
  <img src="https://contrib.rocks/image?repo=ao-org/argentum-online-server" />
</a>




