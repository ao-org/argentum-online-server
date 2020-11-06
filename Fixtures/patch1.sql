ALTER TABLE `ao_server_test`.`user` DROP COLUMN `is_logged`;

CREATE TABLE `ao_server_test`.`statistics` (
  `name` VARCHAR(50) NOT NULL,
  `value` VARCHAR(50) NULL,
  PRIMARY KEY (`name`))
ENGINE = InnoDB;

INSERT INTO `ao_server_test`.`statistics` (`name`, `value`) VALUES ('online', '0');