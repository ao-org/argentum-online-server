# re20-server
Código fuente de el servidor de Argentum20, basado en RevolucionAO de Ladder

Actualizar el archivo Server.ini.produccion con el md5 del cliente. Se obtiene de aqui https://parches.ao20.com.ar/files/Version.json

#Development
Website:
https://ao20-web-testing.herokuapp.com/

#Staging

Website:
https://ao20-web-staging.herokuapp.com/



# Actualizacion

### 1- Crear tag del repositorio de Cliente y Recursos

- https://github.com/ao-org/re20-cliente/releases/new
- https://github.com/ao-org/Recursos/releases/new

Importante usar semantic versioning (https://semver.org/) (ejemplo: v1.0.0)

### 2- Ejecutar pipelines de jenkins para generar parche de cliente. 
Importante: Ejectuar uno solo a la vez y esperar a que termine y hacerlo en este orden.

a- Actualizar Recursos

- http://ao20-test.duckdns.org:9090/view/Produccion/job/Recursos/

b- Actualizar codigo del cliente (ESTE PIPELINE SE EJECUTA AUTOMATICAMENTE AL CREAR EL TAG)

- http://ao20-test.duckdns.org:9090/view/Produccion/job/re20-cliente-tag-release/

c- Crear parche e instalador y subirlo al ftp.

- http://ao20-test.duckdns.org:9090/view/Produccion/job/CREAR%20PARCHE%20CLIENTE%20E%20INSTALADOR%20NUEVO%20PARA%20LA%20WEB/

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
Importante: Ejectuar uno solo a la vez y esperar a que termine y hacerlo en este orden.

- http://ao20-test.duckdns.org:9090/view/Produccion/job/Recursos-tag-release/
- http://ao20-test.duckdns.org:9090/view/Produccion/job/re20-server-tag-release/


Listo parche completado.


