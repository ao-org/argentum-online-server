-- phpMyAdmin SQL Dump
-- version 4.6.6deb5ubuntu0.5
-- https://www.phpmyadmin.net/
--
-- Host: localhost:3306
-- Generation Time: Aug 20, 2021 at 04:41 AM
-- Server version: 5.7.35-0ubuntu0.18.04.1
-- PHP Version: 7.2.24-0ubuntu0.18.04.8

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `ao_server_prod`
--

-- --------------------------------------------------------

--
-- Table structure for table `account`
--

CREATE TABLE `account` (
  `id` int(11) NOT NULL COMMENT 'TRIAL',
  `email` varchar(320) NOT NULL COMMENT 'TRIAL',
  `password` char(64) NOT NULL COMMENT 'TRIAL',
  `salt` char(32) NOT NULL COMMENT 'TRIAL',
  `date_created` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'TRIAL',
  `deleted` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `validated` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `validate_code` char(32) NOT NULL COMMENT 'TRIAL',
  `recovery_code` varchar(32) NOT NULL COMMENT 'TRIAL',
  `is_banned` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `banned_by` varchar(30) NOT NULL COMMENT 'TRIAL',
  `ban_reason` varchar(255) DEFAULT NULL COMMENT 'TRIAL',
  `credits` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `is_donor` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `donor_expire` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'TRIAL',
  `credits_used` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `donor_purchases` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `last_access` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'TRIAL',
  `last_ip` varchar(16) DEFAULT NULL COMMENT 'TRIAL',
  `trial982` char(1) DEFAULT NULL COMMENT 'TRIAL'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='TRIAL';

-- --------------------------------------------------------

--
-- Table structure for table `attribute`
--

CREATE TABLE `attribute` (
  `user_id` int(11) NOT NULL COMMENT 'TRIAL',
  `strength` int(11) NOT NULL COMMENT 'TRIAL',
  `agility` int(11) NOT NULL COMMENT 'TRIAL',
  `intelligence` int(11) NOT NULL COMMENT 'TRIAL',
  `constitution` int(11) NOT NULL COMMENT 'TRIAL',
  `charisma` int(11) NOT NULL COMMENT 'TRIAL',
  `trial989` char(1) DEFAULT NULL COMMENT 'TRIAL'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='TRIAL';

-- --------------------------------------------------------

--
-- Table structure for table `bank_item`
--

CREATE TABLE `bank_item` (
  `user_id` int(11) NOT NULL COMMENT 'TRIAL',
  `number` int(11) NOT NULL COMMENT 'TRIAL',
  `item_id` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `amount` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `trial995` char(1) DEFAULT NULL COMMENT 'TRIAL'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='TRIAL';

-- --------------------------------------------------------

--
-- Table structure for table `house_key`
--

CREATE TABLE `house_key` (
  `key_obj` int(11) NOT NULL COMMENT 'TRIAL',
  `account_id` int(11) NOT NULL COMMENT 'TRIAL',
  `assigned_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'TRIAL',
  `trial002` char(1) DEFAULT NULL COMMENT 'TRIAL'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='TRIAL';

-- --------------------------------------------------------

--
-- Table structure for table `inventory_item`
--

CREATE TABLE `inventory_item` (
  `user_id` int(11) NOT NULL COMMENT 'TRIAL',
  `number` int(11) NOT NULL COMMENT 'TRIAL',
  `item_id` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `amount` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `is_equipped` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `trial005` char(1) DEFAULT NULL COMMENT 'TRIAL'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='TRIAL';

-- --------------------------------------------------------

--
-- Table structure for table `object`
--

CREATE TABLE `object` (
  `number` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `name` varchar(45) DEFAULT NULL COMMENT 'TRIAL',
  `trial015` char(1) DEFAULT NULL COMMENT 'TRIAL'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='TRIAL';

-- --------------------------------------------------------

--
-- Table structure for table `pet`
--

CREATE TABLE `pet` (
  `user_id` int(11) NOT NULL COMMENT 'TRIAL',
  `number` int(11) NOT NULL COMMENT 'TRIAL',
  `pet_id` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `trial015` char(1) DEFAULT NULL COMMENT 'TRIAL'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='TRIAL';

-- --------------------------------------------------------

--
-- Table structure for table `punishment`
--

CREATE TABLE `punishment` (
  `user_id` int(11) NOT NULL COMMENT 'TRIAL',
  `number` int(11) NOT NULL COMMENT 'TRIAL',
  `reason` varchar(255) NOT NULL COMMENT 'TRIAL',
  `created_at` datetime DEFAULT CURRENT_TIMESTAMP COMMENT 'TRIAL',
  `trial018` char(1) DEFAULT NULL COMMENT 'TRIAL'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='TRIAL';

-- --------------------------------------------------------

--
-- Table structure for table `quest`
--

CREATE TABLE `quest` (
  `user_id` int(11) NOT NULL COMMENT 'TRIAL',
  `number` int(11) NOT NULL COMMENT 'TRIAL',
  `quest_id` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `npcs` varchar(64) NOT NULL COMMENT 'TRIAL',
  `npcstarget` varchar(64) NOT NULL COMMENT 'TRIAL',
  `trial021` char(1) DEFAULT NULL COMMENT 'TRIAL'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='TRIAL';

-- --------------------------------------------------------

--
-- Table structure for table `quest_done`
--

CREATE TABLE `quest_done` (
  `user_id` int(11) NOT NULL COMMENT 'TRIAL',
  `quest_id` int(11) NOT NULL COMMENT 'TRIAL',
  `trial028` char(1) DEFAULT NULL COMMENT 'TRIAL'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='TRIAL';

-- --------------------------------------------------------

--
-- Table structure for table `skillpoint`
--

CREATE TABLE `skillpoint` (
  `user_id` int(11) NOT NULL COMMENT 'TRIAL',
  `number` int(11) NOT NULL COMMENT 'TRIAL',
  `value` int(11) NOT NULL COMMENT 'TRIAL',
  `trial034` char(1) DEFAULT NULL COMMENT 'TRIAL'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='TRIAL';

-- --------------------------------------------------------

--
-- Table structure for table `spell`
--

CREATE TABLE `spell` (
  `user_id` int(11) NOT NULL COMMENT 'TRIAL',
  `number` int(11) NOT NULL COMMENT 'TRIAL',
  `spell_id` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `trial041` char(1) DEFAULT NULL COMMENT 'TRIAL'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='TRIAL';

-- --------------------------------------------------------

--
-- Table structure for table `statistics`
--

CREATE TABLE `statistics` (
  `name` varchar(50) NOT NULL COMMENT 'TRIAL',
  `value` varchar(50) DEFAULT NULL COMMENT 'TRIAL',
  `trial051` char(1) DEFAULT NULL COMMENT 'TRIAL'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='TRIAL';

-- --------------------------------------------------------

--
-- Table structure for table `user`
--

CREATE TABLE `user` (
  `id` int(11) NOT NULL COMMENT 'TRIAL',
  `account_id` int(11) NOT NULL COMMENT 'TRIAL',
  `deleted` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `name` varchar(30) NOT NULL COMMENT 'TRIAL',
  `level` int(11) NOT NULL COMMENT 'TRIAL',
  `exp` int(11) NOT NULL COMMENT 'TRIAL',
  `genre_id` int(11) NOT NULL COMMENT 'TRIAL',
  `race_id` int(11) NOT NULL COMMENT 'TRIAL',
  `class_id` int(11) NOT NULL COMMENT 'TRIAL',
  `home_id` int(11) NOT NULL COMMENT 'TRIAL',
  `description` varchar(255) DEFAULT NULL COMMENT 'TRIAL',
  `gold` int(11) NOT NULL COMMENT 'TRIAL',
  `bank_gold` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `free_skillpoints` int(11) NOT NULL COMMENT 'TRIAL',
  `pets_saved` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `votes_amount` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `spouse` varchar(30) NOT NULL COMMENT 'TRIAL',
  `message_info` varchar(512) DEFAULT NULL COMMENT 'TRIAL',
  `pos_map` int(11) NOT NULL COMMENT 'TRIAL',
  `pos_x` int(11) NOT NULL COMMENT 'TRIAL',
  `pos_y` int(11) NOT NULL COMMENT 'TRIAL',
  `body_id` int(11) NOT NULL COMMENT 'TRIAL',
  `head_id` int(11) NOT NULL COMMENT 'TRIAL',
  `weapon_id` int(11) NOT NULL COMMENT 'TRIAL',
  `helmet_id` int(11) NOT NULL COMMENT 'TRIAL',
  `shield_id` int(11) NOT NULL COMMENT 'TRIAL',
  `heading` int(11) NOT NULL DEFAULT '3' COMMENT 'TRIAL',
  `slot_armour` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_weapon` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_helmet` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_shield` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_ammo` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_ship` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_mount` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_dm` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_rm` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_magic` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_knuckles` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_tool` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `min_hp` int(11) NOT NULL COMMENT 'TRIAL',
  `max_hp` int(11) NOT NULL COMMENT 'TRIAL',
  `min_man` int(11) NOT NULL COMMENT 'TRIAL',
  `max_man` int(11) NOT NULL COMMENT 'TRIAL',
  `min_sta` int(11) NOT NULL COMMENT 'TRIAL',
  `max_sta` int(11) NOT NULL COMMENT 'TRIAL',
  `min_ham` int(11) NOT NULL COMMENT 'TRIAL',
  `max_ham` int(11) NOT NULL COMMENT 'TRIAL',
  `min_sed` int(11) NOT NULL COMMENT 'TRIAL',
  `max_sed` int(11) NOT NULL COMMENT 'TRIAL',
  `min_hit` int(11) NOT NULL COMMENT 'TRIAL',
  `max_hit` int(11) NOT NULL COMMENT 'TRIAL',
  `killed_npcs` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `killed_users` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `invent_level` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_naked` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_poisoned` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_incinerated` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_dead` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_sailing` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_paralyzed` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_silenced` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `silence_minutes_left` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `silence_elapsed_seconds` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `is_mounted` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_banned` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `banned_by` varchar(30) NOT NULL COMMENT 'TRIAL',
  `ban_reason` varchar(255) DEFAULT NULL COMMENT 'TRIAL',
  `counter_pena` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `deaths` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `pertenece_consejo_real` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `pertenece_consejo_caos` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `pertenece_real` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `pertenece_caos` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `ciudadanos_matados` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `criminales_matados` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `recibio_armadura_real` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `recibio_armadura_caos` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `recibio_exp_real` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `recibio_exp_caos` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `recompensas_real` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `recompensas_caos` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `reenlistadas` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `fecha_ingreso` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'TRIAL',
  `nivel_ingreso` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `matados_ingreso` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `siguiente_recompensa` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `status` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `guild_index` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `guild_aspirant_index` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `guild_member_history` varchar(1024) DEFAULT NULL COMMENT 'TRIAL',
  `guild_requests_history` varchar(1024) DEFAULT NULL COMMENT 'TRIAL',
  `guild_rejected_because` varchar(255) DEFAULT NULL COMMENT 'TRIAL',
  `chat_global` int(11) DEFAULT '1' COMMENT 'TRIAL',
  `chat_combate` int(11) DEFAULT '1' COMMENT 'TRIAL',
  `warnings` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `elo` int(11) NOT NULL DEFAULT '1000' COMMENT 'TRIAL',
  `return_map` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `return_x` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `return_y` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `last_logout` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `trial051` char(1) DEFAULT NULL COMMENT 'TRIAL'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='TRIAL';

-- --------------------------------------------------------

--
-- Table structure for table `user_deleted`
--

CREATE TABLE `user_deleted` (
  `id` int(11) NOT NULL COMMENT 'TRIAL',
  `account_id` int(11) NOT NULL COMMENT 'TRIAL',
  `deleted` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'TRIAL',
  `name` varchar(30) NOT NULL COMMENT 'TRIAL',
  `level` int(11) NOT NULL COMMENT 'TRIAL',
  `exp` int(11) NOT NULL COMMENT 'TRIAL',
  `genre_id` int(11) NOT NULL COMMENT 'TRIAL',
  `race_id` int(11) NOT NULL COMMENT 'TRIAL',
  `class_id` int(11) NOT NULL COMMENT 'TRIAL',
  `home_id` int(11) NOT NULL COMMENT 'TRIAL',
  `description` varchar(255) DEFAULT NULL COMMENT 'TRIAL',
  `gold` int(11) NOT NULL COMMENT 'TRIAL',
  `bank_gold` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `free_skillpoints` int(11) NOT NULL COMMENT 'TRIAL',
  `pets_saved` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `votes_amount` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `spouse` varchar(30) NOT NULL COMMENT 'TRIAL',
  `message_info` varchar(512) DEFAULT NULL COMMENT 'TRIAL',
  `pos_map` int(11) NOT NULL COMMENT 'TRIAL',
  `pos_x` int(11) NOT NULL COMMENT 'TRIAL',
  `pos_y` int(11) NOT NULL COMMENT 'TRIAL',
  `body_id` int(11) NOT NULL COMMENT 'TRIAL',
  `head_id` int(11) NOT NULL COMMENT 'TRIAL',
  `weapon_id` int(11) NOT NULL COMMENT 'TRIAL',
  `helmet_id` int(11) NOT NULL COMMENT 'TRIAL',
  `shield_id` int(11) NOT NULL COMMENT 'TRIAL',
  `heading` int(11) NOT NULL DEFAULT '3' COMMENT 'TRIAL',
  `slot_armour` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_weapon` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_helmet` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_shield` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_ammo` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_ship` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_mount` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_dm` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_rm` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_magic` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_knuckles` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `slot_tool` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `min_hp` int(11) NOT NULL COMMENT 'TRIAL',
  `max_hp` int(11) NOT NULL COMMENT 'TRIAL',
  `min_man` int(11) NOT NULL COMMENT 'TRIAL',
  `max_man` int(11) NOT NULL COMMENT 'TRIAL',
  `min_sta` int(11) NOT NULL COMMENT 'TRIAL',
  `max_sta` int(11) NOT NULL COMMENT 'TRIAL',
  `min_ham` int(11) NOT NULL COMMENT 'TRIAL',
  `max_ham` int(11) NOT NULL COMMENT 'TRIAL',
  `min_sed` int(11) NOT NULL COMMENT 'TRIAL',
  `max_sed` int(11) NOT NULL COMMENT 'TRIAL',
  `min_hit` int(11) NOT NULL COMMENT 'TRIAL',
  `max_hit` int(11) NOT NULL COMMENT 'TRIAL',
  `killed_npcs` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `killed_users` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `invent_level` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_naked` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_poisoned` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_incinerated` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_dead` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_sailing` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_paralyzed` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_logged` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_silenced` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `silence_minutes_left` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `silence_elapsed_seconds` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `is_mounted` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `is_banned` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `banned_by` varchar(30) NOT NULL COMMENT 'TRIAL',
  `ban_reason` varchar(255) DEFAULT NULL COMMENT 'TRIAL',
  `counter_pena` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `deaths` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `pertenece_consejo_real` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `pertenece_consejo_caos` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `pertenece_real` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `pertenece_caos` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `ciudadanos_matados` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `criminales_matados` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `recibio_armadura_real` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `recibio_armadura_caos` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `recibio_exp_real` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `recibio_exp_caos` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `recompensas_real` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `recompensas_caos` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `reenlistadas` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `fecha_ingreso` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'TRIAL',
  `nivel_ingreso` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `matados_ingreso` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `siguiente_recompensa` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `status` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `guild_index` int(11) DEFAULT '0' COMMENT 'TRIAL',
  `guild_aspirant_index` int(11) DEFAULT NULL COMMENT 'TRIAL',
  `guild_member_history` varchar(1024) DEFAULT NULL COMMENT 'TRIAL',
  `guild_requests_history` varchar(1024) DEFAULT NULL COMMENT 'TRIAL',
  `guild_rejected_because` varchar(255) DEFAULT NULL COMMENT 'TRIAL',
  `chat_global` int(11) DEFAULT '1' COMMENT 'TRIAL',
  `chat_combate` int(11) DEFAULT '1' COMMENT 'TRIAL',
  `warnings` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `elo` int(11) NOT NULL DEFAULT '1000' COMMENT 'TRIAL',
  `return_map` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `return_x` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `return_y` int(11) NOT NULL DEFAULT '0' COMMENT 'TRIAL',
  `trial074` char(1) DEFAULT NULL COMMENT 'TRIAL'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='TRIAL';

--
-- Indexes for dumped tables
--

--
-- Indexes for table `account`
--
ALTER TABLE `account`
  ADD PRIMARY KEY (`id`),
  ADD KEY `idx_delete_account` (`deleted`),
  ADD KEY `idx_account_email` (`email`(255));

--
-- Indexes for table `attribute`
--
ALTER TABLE `attribute`
  ADD PRIMARY KEY (`user_id`);

--
-- Indexes for table `bank_item`
--
ALTER TABLE `bank_item`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indexes for table `house_key`
--
ALTER TABLE `house_key`
  ADD PRIMARY KEY (`key_obj`),
  ADD KEY `idx_house_key_fk_account` (`account_id`);

--
-- Indexes for table `inventory_item`
--
ALTER TABLE `inventory_item`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indexes for table `pet`
--
ALTER TABLE `pet`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indexes for table `punishment`
--
ALTER TABLE `punishment`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indexes for table `quest`
--
ALTER TABLE `quest`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indexes for table `quest_done`
--
ALTER TABLE `quest_done`
  ADD PRIMARY KEY (`user_id`,`quest_id`);

--
-- Indexes for table `skillpoint`
--
ALTER TABLE `skillpoint`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indexes for table `spell`
--
ALTER TABLE `spell`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indexes for table `statistics`
--
ALTER TABLE `statistics`
  ADD PRIMARY KEY (`name`);

--
-- Indexes for table `user`
--
ALTER TABLE `user`
  ADD PRIMARY KEY (`id`,`account_id`,`deleted`),
  ADD KEY `idx_user_name` (`name`),
  ADD KEY `idx_user_fk_user_account` (`account_id`);

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `account`
--
ALTER TABLE `account`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT COMMENT 'TRIAL', AUTO_INCREMENT=296190;
--
-- AUTO_INCREMENT for table `attribute`
--
ALTER TABLE `attribute`
  MODIFY `user_id` int(11) NOT NULL AUTO_INCREMENT COMMENT 'TRIAL', AUTO_INCREMENT=1562737;
--
-- AUTO_INCREMENT for table `bank_item`
--
ALTER TABLE `bank_item`
  MODIFY `user_id` int(11) NOT NULL AUTO_INCREMENT COMMENT 'TRIAL', AUTO_INCREMENT=2066;
--
-- AUTO_INCREMENT for table `house_key`
--
ALTER TABLE `house_key`
  MODIFY `key_obj` int(11) NOT NULL AUTO_INCREMENT COMMENT 'TRIAL', AUTO_INCREMENT=344;
--
-- AUTO_INCREMENT for table `inventory_item`
--
ALTER TABLE `inventory_item`
  MODIFY `user_id` int(11) NOT NULL AUTO_INCREMENT COMMENT 'TRIAL', AUTO_INCREMENT=2304;
--
-- AUTO_INCREMENT for table `pet`
--
ALTER TABLE `pet`
  MODIFY `user_id` int(11) NOT NULL AUTO_INCREMENT COMMENT 'TRIAL', AUTO_INCREMENT=1562740;
--
-- AUTO_INCREMENT for table `punishment`
--
ALTER TABLE `punishment`
  MODIFY `user_id` int(11) NOT NULL AUTO_INCREMENT COMMENT 'TRIAL', AUTO_INCREMENT=13227;
--
-- AUTO_INCREMENT for table `quest`
--
ALTER TABLE `quest`
  MODIFY `user_id` int(11) NOT NULL AUTO_INCREMENT COMMENT 'TRIAL', AUTO_INCREMENT=1562742;
--
-- AUTO_INCREMENT for table `quest_done`
--
ALTER TABLE `quest_done`
  MODIFY `user_id` int(11) NOT NULL AUTO_INCREMENT COMMENT 'TRIAL', AUTO_INCREMENT=13011;
--
-- AUTO_INCREMENT for table `skillpoint`
--
ALTER TABLE `skillpoint`
  MODIFY `user_id` int(11) NOT NULL AUTO_INCREMENT COMMENT 'TRIAL', AUTO_INCREMENT=3584;
--
-- AUTO_INCREMENT for table `spell`
--
ALTER TABLE `spell`
  MODIFY `user_id` int(11) NOT NULL AUTO_INCREMENT COMMENT 'TRIAL', AUTO_INCREMENT=3280;
--
-- AUTO_INCREMENT for table `user`
--
ALTER TABLE `user`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT COMMENT 'TRIAL', AUTO_INCREMENT=13247;
--
-- Constraints for dumped tables
--

--
-- Constraints for table `bank_item`
--
ALTER TABLE `bank_item`
  ADD CONSTRAINT `fk_attribute_user_0` FOREIGN KEY (`user_id`) REFERENCES `user` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Constraints for table `house_key`
--
ALTER TABLE `house_key`
  ADD CONSTRAINT `fk_bank_item_account_0` FOREIGN KEY (`account_id`) REFERENCES `account` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Constraints for table `inventory_item`
--
ALTER TABLE `inventory_item`
  ADD CONSTRAINT `fk_house_key_user_0` FOREIGN KEY (`user_id`) REFERENCES `user` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Constraints for table `spell`
--
ALTER TABLE `spell`
  ADD CONSTRAINT `fk_skillpoint_user_0` FOREIGN KEY (`user_id`) REFERENCES `user` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Constraints for table `user`
--
ALTER TABLE `user`
  ADD CONSTRAINT `fk_statistics_account_0` FOREIGN KEY (`account_id`) REFERENCES `account` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
