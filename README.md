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


# Development (test branch ao-api/web)
IMPORTANTE: Para hacer cuentas en el servidor de test, tienen que entrar aca.
Website:
http://test.ao20.com.ar



# Staging (test master ao-api/web)
IMPORTANTE: Para hacer cuentas en el servidor de staging, tienen que entrar aca.
Website:
http://staging.ao20.com.ar

# Creacion de Parches / Actualizacion

### 1- Crear tag del repositorio de Cliente y Recursos

- https://github.com/ao-org/re20-cliente/releases/new
- https://github.com/ao-org/Recursos/releases/new

Importante usar semantic versioning (https://semver.org/) (ejemplo: v1.0.0)

### 2- Ejecutar pipelines de jenkins para generar parche de cliente. 
Importante: Ejecutar uno solo a la vez y esperar a que termine y hacerlo en este orden.

a- Actualizar Recursos

- http://horacio.ao20.com.ar:2095/view/Produccion/job/Recursos-tag-release-cliente/

b- Actualizar codigo del cliente (ESTE PIPELINE SE EJECUTA AUTOMATICAMENTE AL CREAR EL TAG)

- http://horacio.ao20.com.ar:2095/view/Produccion/job/re20-cliente-tag-release/

c- Crear parche e instalador y subirlo al ftp.

- http://horacio.ao20.com.ar:2095/view/Produccion/job/CREAR%20PARCHE%20CLIENTE%20E%20INSTALADOR%20NUEVO%20PARA%20LA%20WEB/

Esperar a que termine.
PRO TIP: Se puede continuar el proceso cuando se esta generando/subiendo el instalador, ya que el parche estaria completo, si no sabes bien cuando esto sucede, simplemente esperar a que termine el proceso.

### 2- Forzar actualizacion cliente y recursos a usuarios.
Entrar al VPS y hacer click en el boton de server que dice: `Cerrar server y forzar actualizacion` para forzar actualizar los recursos y cliente a los usuarios.

### 3- Poner MD5 de Cliente en Server.ini
Una vez finalizado, ingresar a: https://parches.ao20.com.ar/files/Version.json y copiar el `md5` del cliente (`Argentum20\Cliente\Argentum.exe`)
y pegarlo en el Server.ini.Produccion (https://github.com/ao-org/re20-server/blob/master/Server.ini.produccion) en la propiedad `[CHECKSUM] -> Cliente` y commitear.

### 4- Crear tag del repositorio del server 

- https://github.com/ao-org/re20-server/releases/new

Importante usar semantic versioning (https://semver.org/) (ejemplo: v1.0.0)

### 5- Ejecutar pipelines de jenkins para actualizar recursos de servidor y servidor.
Importante: Ejecutar uno solo a la vez y esperar a que termine y hacerlo en este orden.

- http://horacio.ao20.com.ar:2095/view/Produccion/job/Recursos-tag-release/
- http://horacio.ao20.com.ar:2095/view/Produccion/job/re20-server-tag-release/


Listo parche completado.


