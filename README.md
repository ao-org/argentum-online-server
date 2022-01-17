# re20-server
Código fuente de el servidor de Argentum20, basado en RevolucionAO de Ladder

![image](https://user-images.githubusercontent.com/5874806/126402326-e94f25b3-3992-4db2-ad0b-8b30ad5d34ee.png)

# Requisitos

- ODBC DRIVERS
Mysql
https://dev.mysql.com/get/Downloads/Connector-ODBC/8.0/mysql-connector-odbc-8.0.22-win32.msi

SQlite
http://www.ch-werner.de/sqliteodbc/

- Registrar manualmente libreria Aurora.Network.dll 

Abrir CMD como Administrador `regsvr32 Aurora.Network.dll`

# Staging (test master ao-api/web)
IMPORTANTE: Para hacer cuentas en el servidor de staging, tienen que entrar aca.
Website:
http://staging.ao20.com.ar

# Creacion de Parches / Actualizacion

### 1- Crear tag del repositorio de Cliente, Server y Recursos

- https://github.com/ao-org/re20-cliente/releases/new
- https://github.com/ao-org/Recursos/releases/new
- https://github.com/ao-org/re20-server/releases/new


Importante usar semantic versioning (https://semver.org/) (ejemplo: v1.0.0)

### 2- Ejecutar pipeline de jenkins para generar parche completo (server, cliente y recursos). 

http://horacio.ao20.com.ar:2095/job/1%20-%20CREAR%20PARCHE%20TOTAL/

Listo parche completado.


