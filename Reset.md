## Reset de server:

Cambio archivo Configuracion.ini en repositorio de argentum20-server con exp y oro x1 ☑️ 

Borrar viejos backup FTP ☑️ 

ToggleFeatures ☑️ 

RecordUsers.log puesto en 0 ☑️ 

NumUsers.log puesto en 0 ☑️ 

Borrar contenido de GenteBanned.log y BanDetailed.dat  ☑️ 

Borrar errores.log (los que hayan en carpeta logs) ☑️ 

Configuracion de Pymmo para habilitar entrada de solo usuarios Patreon ☑️


### Hay que correr la siguiente query, esto no borra las cuentas de SQLITE.
- Database.db
Utilizar Empty_Database.db y migrarla con comando `argentums20-assets\tools\MigratePatreonAccounts.exe Database.db Empty_Database.db`

- ao20_gameserver_clone (MySql)
```
SET @OLD_FOREIGN_KEY_CHECKS = @@FOREIGN_KEY_CHECKS;
SET FOREIGN_KEY_CHECKS = 0;

TRUNCATE TABLE `bank_item`;
TRUNCATE TABLE `char_transfer_record`;
TRUNCATE TABLE `guild_members`;
TRUNCATE TABLE `guild_request`;
TRUNCATE TABLE `guild_member_history`;
TRUNCATE TABLE `guild_request_history`;
-- DELETE statement instead of TRUNCATE for `guilds`
DELETE FROM `guilds`;
TRUNCATE TABLE `inventory_item`;
TRUNCATE TABLE `patreon_shop_audit`;
TRUNCATE TABLE `pet`;
TRUNCATE TABLE `quest`;
TRUNCATE TABLE `quest_done`;
TRUNCATE TABLE `skillpoint`;
TRUNCATE TABLE `spell`;
TRUNCATE TABLE `statistics_users_online`;
TRUNCATE TABLE `tokens`;
-- DELETE statement instead of TRUNCATE for `user`
DELETE FROM `user`;

SET FOREIGN_KEY_CHECKS = @OLD_FOREIGN_KEY_CHECKS;
```

- ao20 sqlite
```
DELETE FROM bank_item;
DELETE FROM char_transfer_record;
DELETE FROM guilds;
DELETE FROM guild_members;
DELETE FROM guild_member_history;
DELETE FROM guild_request;
DELETE FROM guild_request_history;
DELETE FROM inventory_item;
DELETE FROM patreon_shop_audit;
DELETE FROM pet;
DELETE FROM quest;
DELETE FROM quest_done;
DELETE FROM skillpoint;
DELETE FROM spell;
DELETE FROM tokens;
DELETE FROM user;
DELETE FROM punishment;
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

