## Reset de server:

Cambio archivo Configuracion.ini en repositorio de argentum20-server con exp y oro x1 ☑️ 

Borrar Clanes y dejar guilds.inf default☑️ 

Borrar Clanes FTP ☑️ 

Borrar viejos backup FTP ☑️ 

ToggleFeatures ☑️ 

RecordUsers.log puesto en 0 ☑️ 

NumUsers.log puesto en 0 ☑️ 

Borrar contenido de GenteBanned.log y BanDetailed.dat  ☑️ 

Borrar errores.log (los que hayan en carpeta logs) ☑️ 


### Hay que correr la siguiente query, esto no borra las cuentas de SQLITE.
- Database.db
Utilizar Empty_Database.db y migrarla con comando `argentums20-assets\tools\MigratePatreonAccounts.exe Database.db Empty_Database.db`

- ao20_gameserver_clone (MySql)
```
SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0;
TRUNCATE TABLE `account`;
TRUNCATE TABLE `attribute`;
TRUNCATE TABLE `bank_item`;
TRUNCATE TABLE `char_transfer_record`;
TRUNCATE TABLE `inventory_item`;
TRUNCATE TABLE `patreon_shop_audit`;
TRUNCATE TABLE `pet`;
TRUNCATE TABLE `quest`;
TRUNCATE TABLE `quest_done`;
TRUNCATE TABLE `skillpoint`;
TRUNCATE TABLE `spell`;
TRUNCATE TABLE `statistics_users_online`;
TRUNCATE TABLE `tokens`;
TRUNCATE TABLE `user`;
TRUNCATE TABLE `mercadopago_account`;
SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS;
```

- ao20_pymmo
```
SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0;
TRUNCATE TABLE `all_house_keys`;
TRUNCATE TABLE `gold_statistics`;
TRUNCATE TABLE `items_statistics`;
TRUNCATE TABLE `patron_pc_ranking`;
SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS;
```

