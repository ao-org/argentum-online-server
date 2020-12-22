ALTER TABLE `attribute` DROP FOREIGN KEY `fk_attribute_user`; 
ALTER TABLE `attribute` ADD CONSTRAINT `fk_attribute_user` FOREIGN KEY (`user_id`) REFERENCES `user`(`id`) ON DELETE CASCADE ON UPDATE CASCADE;

ALTER TABLE `bank_item` DROP FOREIGN KEY `fk_bank_user`; 
ALTER TABLE `bank_item` ADD CONSTRAINT `fk_bank_user` FOREIGN KEY (`user_id`) REFERENCES `user`(`id`) ON DELETE CASCADE ON UPDATE CASCADE;

ALTER TABLE `connection` DROP FOREIGN KEY `fk_ip_user`; 
ALTER TABLE `connection` ADD CONSTRAINT `fk_ip_user` FOREIGN KEY (`user_id`) REFERENCES `user`(`id`) ON DELETE CASCADE ON UPDATE CASCADE;

ALTER TABLE `house_key` DROP FOREIGN KEY `fk_account`; 
ALTER TABLE `house_key` ADD CONSTRAINT `fk_account` FOREIGN KEY (`account_id`) REFERENCES `account`(`id`) ON DELETE CASCADE ON UPDATE CASCADE;
	
ALTER TABLE `inventory_item` DROP FOREIGN KEY `fk_inventory_user`; 
ALTER TABLE `inventory_item` ADD CONSTRAINT `fk_inventory_user` FOREIGN KEY (`user_id`) REFERENCES `user`(`id`) ON DELETE CASCADE ON UPDATE CASCADE;

ALTER TABLE `mail` CHANGE `date` `date` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP;
ALTER TABLE `mail` ADD `id` MEDIUMINT(8) UNSIGNED NOT NULL AUTO_INCREMENT FIRST, ADD PRIMARY KEY (`id`);
ALTER TABLE `mail` DROP FOREIGN KEY `fk_mail`; 
ALTER TABLE `mail` ADD CONSTRAINT `fk_mail` FOREIGN KEY (`user_id`) REFERENCES `user`(`id`) ON DELETE CASCADE ON UPDATE CASCADE;

ALTER TABLE `pet` DROP FOREIGN KEY `fk_pet_user`; 
ALTER TABLE `pet` ADD CONSTRAINT `fk_pet_user` FOREIGN KEY (`user_id`) REFERENCES `user`(`id`) ON DELETE CASCADE ON UPDATE CASCADE;

ALTER TABLE `punishment` DROP FOREIGN KEY `fk_punishment_user`; 
ALTER TABLE `punishment` ADD CONSTRAINT `fk_punishment_user` FOREIGN KEY (`user_id`) REFERENCES `user`(`id`) ON DELETE CASCADE ON UPDATE CASCADE;

ALTER TABLE `quest` DROP FOREIGN KEY `fk_quest_user`; 
ALTER TABLE `quest` ADD CONSTRAINT `fk_quest_user` FOREIGN KEY (`user_id`) REFERENCES `user`(`id`) ON DELETE CASCADE ON UPDATE CASCADE;

ALTER TABLE `quest_done` DROP FOREIGN KEY `fk_quest_done_user`; 
ALTER TABLE `quest_done` ADD CONSTRAINT `fk_quest_done_user` FOREIGN KEY (`user_id`) REFERENCES `user`(`id`) ON DELETE CASCADE ON UPDATE CASCADE;

ALTER TABLE `skillpoint` DROP FOREIGN KEY `fk_skillpoint_user`;
ALTER TABLE `skillpoint` ADD CONSTRAINT `fk_skillpoint_user` FOREIGN KEY (`user_id`) REFERENCES `user`(`id`) ON DELETE CASCADE ON UPDATE CASCADE;

ALTER TABLE `spell` DROP FOREIGN KEY `fk_spell_user`; 
ALTER TABLE `spell` ADD CONSTRAINT `fk_spell_user` FOREIGN KEY (`user_id`) REFERENCES `user`(`id`) ON DELETE CASCADE ON UPDATE CASCADE;

ALTER TABLE `user` DROP FOREIGN KEY `fk_user_account`; 
ALTER TABLE `user` ADD CONSTRAINT `fk_user_account` FOREIGN KEY (`account_id`) REFERENCES `account`(`id`) ON DELETE CASCADE ON UPDATE CASCADE;