ALTER TABLE `user` 
ADD COLUMN `pets_saved` TINYINT(1) NOT NULL DEFAULT '0' AFTER `pet_amount`;