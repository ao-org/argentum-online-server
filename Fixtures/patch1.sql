CREATE TABLE `statistics` (
  `name` VARCHAR(50) NOT NULL,
  `value` VARCHAR(50) NULL,
  PRIMARY KEY (`name`))
ENGINE = InnoDB;

INSERT INTO `statistics` (`name`, `value`) VALUES ('online', '0');

ALTER TABLE `account` CHANGE COLUMN `is_logged` `logged` INT(11) NULL DEFAULT 0;