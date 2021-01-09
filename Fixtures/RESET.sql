-- Comentadas debido a que normalmente no se resetan las cuentas y sus tablas relacionadas.
-- TRUNCATE TABLE ao_server_prod.account;
-- TRUNCATE TABLE ao_server_prod.house_key;

TRUNCATE TABLE ao_server_prod.attribute;
TRUNCATE TABLE ao_server_prod.bank_item;
TRUNCATE TABLE ao_server_prod.connection;
TRUNCATE TABLE ao_server_prod.inventory_item;
TRUNCATE TABLE ao_server_prod.mail;
TRUNCATE TABLE ao_server_prod.pet;
TRUNCATE TABLE ao_server_prod.punishment;
TRUNCATE TABLE ao_server_prod.quest;
TRUNCATE TABLE ao_server_prod.quest_done;
TRUNCATE TABLE ao_server_prod.skillpoint;
TRUNCATE TABLE ao_server_prod.spell;

-- En vez de vaciar la tabla, seteamos los valores a 0 para no romper la Web.
REPLACE INTO ao_server_prod.statistics (name, value) VALUES ('online', '0');
REPLACE INTO ao_server_prod.statistics (name, value) VALUES ('record', '0');

TRUNCATE TABLE ao_server_prod.user;
