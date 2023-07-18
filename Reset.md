### Para resetear el servidor hay que correr la siguiente query, esto no borra las cuentas.

- Database.db (Sqlite) & ao20_gameserver_clone (MySql)
```
SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0;
DROP TABLE `attribute`;
DROP TABLE `bank_item`;
DROP TABLE `char_transfer_record`;
DROP TABLE `inventory_item`;
DROP TABLE `patreon_shop_audit`;
DROP TABLE `pet`;
DROP TABLE `quest`;
DROP TABLE `quest_done`;
DROP TABLE `skillpoint`;
DROP TABLE `spell`;
DROP TABLE `statistics_users_online`;
DROP TABLE `tokens`;
DROP TABLE `user`;
SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS;
```

- ao20_pymmo
```
SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0;
DROP TABLE `all_house_keys`;
DROP TABLE `gold_statistics`;
DROP TABLE `items_statistics`;
DROP TABLE `patron_pc_ranking`;
SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS;
```
