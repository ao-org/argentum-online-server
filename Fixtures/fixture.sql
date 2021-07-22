-- phpMyAdmin SQL Dump
-- version 5.1.0
-- https://www.phpmyadmin.net/
--
-- Servidor: 127.0.0.1
-- Tiempo de generación: 23-07-2021 a las 01:07:42
-- Versión del servidor: 10.4.18-MariaDB
-- Versión de PHP: 8.0.3

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Base de datos: `ao_server_prod`
--

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `account`
--

CREATE TABLE `account` (
  `id` mediumint(8) UNSIGNED NOT NULL,
  `email` varchar(320) NOT NULL,
  `password` char(64) NOT NULL,
  `salt` char(32) NOT NULL,
  `date_created` timestamp NOT NULL DEFAULT current_timestamp(),
  `mac_address` char(17) DEFAULT '',
  `hd_serial` int(11) NOT NULL DEFAULT 0,
  `logged` int(11) NOT NULL DEFAULT 0,
  `deleted` tinyint(1) DEFAULT 0,
  `validated` tinyint(1) NOT NULL DEFAULT 0,
  `validate_code` char(32) NOT NULL,
  `recovery_code` varchar(32) NOT NULL DEFAULT '',
  `is_banned` tinyint(1) DEFAULT 0,
  `banned_by` varchar(30) NOT NULL DEFAULT '',
  `ban_reason` varchar(255) DEFAULT '',
  `credits` int(10) UNSIGNED DEFAULT 0,
  `is_donor` tinyint(1) DEFAULT 0,
  `donor_expire` timestamp NOT NULL DEFAULT current_timestamp(),
  `credits_used` int(10) UNSIGNED DEFAULT 0,
  `donor_purchases` int(10) UNSIGNED DEFAULT 0,
  `last_access` timestamp NOT NULL DEFAULT current_timestamp(),
  `last_ip` varchar(16) DEFAULT ''
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `attribute`
--

CREATE TABLE `attribute` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `value` tinyint(3) UNSIGNED NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `bank_item`
--

CREATE TABLE `bank_item` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `item_id` smallint(5) UNSIGNED DEFAULT NULL,
  `amount` smallint(5) UNSIGNED DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `connection`
--

CREATE TABLE `connection` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `ip` varchar(16) NOT NULL,
  `date_last_login` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `house_key`
--

CREATE TABLE `house_key` (
  `key_obj` smallint(5) UNSIGNED NOT NULL,
  `account_id` mediumint(8) UNSIGNED NOT NULL,
  `assigned_at` timestamp NOT NULL DEFAULT current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `inventory_item`
--

CREATE TABLE `inventory_item` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `item_id` smallint(5) UNSIGNED DEFAULT NULL,
  `amount` smallint(5) UNSIGNED DEFAULT NULL,
  `is_equipped` tinyint(1) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `mail`
--

CREATE TABLE `mail` (
  `id` mediumint(8) UNSIGNED NOT NULL,
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `sender` varchar(30) NOT NULL DEFAULT '',
  `item_id` smallint(5) UNSIGNED DEFAULT 0,
  `amount` smallint(5) UNSIGNED DEFAULT 0,
  `date` timestamp NOT NULL DEFAULT current_timestamp(),
  `is_read` tinyint(1) DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `mao__claims`
--

CREATE TABLE `mao__claims` (
  `id` int(11) UNSIGNED NOT NULL,
  `deal_id` int(11) UNSIGNED NOT NULL,
  `claimer_account_id` mediumint(8) UNSIGNED NOT NULL,
  `claimer_character_id` mediumint(8) UNSIGNED DEFAULT NULL,
  `gold_amount` int(11) NOT NULL,
  `claimed_at` timestamp NULL DEFAULT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `mao__deals`
--

CREATE TABLE `mao__deals` (
  `id` int(11) UNSIGNED NOT NULL,
  `publisher_account_id` mediumint(8) UNSIGNED NOT NULL,
  `character_id` mediumint(8) UNSIGNED NOT NULL,
  `description` varchar(512) DEFAULT '',
  `deal_type` varchar(11) NOT NULL COMMENT 'Any of `sell/trade`',
  `gold_price` int(11) NOT NULL,
  `closed_at` timestamp NULL DEFAULT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura Stand-in para la vista `mao__deal_summaries`
-- (Véase abajo para la vista actual)
--
CREATE TABLE `mao__deal_summaries` (
`id` int(11) unsigned
,`publisher_account_id` mediumint(8) unsigned
,`character_id` mediumint(8) unsigned
,`description` varchar(512)
,`deal_type` varchar(11)
,`gold_price` int(11)
,`closed_at` timestamp
,`created_at` timestamp
,`updated_at` timestamp
,`character_name` varchar(30)
,`class_id` tinyint(3) unsigned
,`race_id` tinyint(3) unsigned
,`genre_id` tinyint(3) unsigned
,`elo` int(11)
,`level` smallint(5) unsigned
,`head_id` smallint(5) unsigned
,`max_hp` smallint(5) unsigned
,`warnings` tinyint(3) unsigned
);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `mao__offers`
--

CREATE TABLE `mao__offers` (
  `id` int(11) UNSIGNED NOT NULL,
  `deal_id` int(11) UNSIGNED NOT NULL,
  `offeror_account_id` mediumint(8) UNSIGNED NOT NULL,
  `character_id` mediumint(8) UNSIGNED NOT NULL,
  `offer_price` int(11) NOT NULL,
  `state` varchar(255) NOT NULL COMMENT 'Any of `pending/accepted/rejected/completed/cancelled`',
  `description` varchar(512) DEFAULT '',
  `created_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura Stand-in para la vista `mao__offer_summaries`
-- (Véase abajo para la vista actual)
--
CREATE TABLE `mao__offer_summaries` (
`id` int(11) unsigned
,`deal_id` int(11) unsigned
,`offeror_account_id` mediumint(8) unsigned
,`character_id` mediumint(8) unsigned
,`offer_price` int(11)
,`state` varchar(255)
,`description` varchar(512)
,`created_at` timestamp
,`updated_at` timestamp
,`character_name` varchar(30)
,`class_id` tinyint(3) unsigned
,`race_id` tinyint(3) unsigned
,`genre_id` tinyint(3) unsigned
,`elo` int(11)
,`level` smallint(5) unsigned
,`head_id` smallint(5) unsigned
,`max_hp` smallint(5) unsigned
,`warnings` tinyint(3) unsigned
);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `pet`
--

CREATE TABLE `pet` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `pet_id` smallint(5) UNSIGNED DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `punishment`
--

CREATE TABLE `punishment` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL DEFAULT 0,
  `reason` varchar(255) NOT NULL,
  `created_at` timestamp NULL DEFAULT current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `quest`
--

CREATE TABLE `quest` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `quest_id` tinyint(3) UNSIGNED NOT NULL DEFAULT 0,
  `npcs` varchar(64) NOT NULL DEFAULT '',
  `npcstarget` varchar(64) NOT NULL DEFAULT ''
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `quest_done`
--

CREATE TABLE `quest_done` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `quest_id` tinyint(3) UNSIGNED NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura Stand-in para la vista `ranking_users`
-- (Véase abajo para la vista actual)
--
CREATE TABLE `ranking_users` (
`id` mediumint(8) unsigned
,`character_name` varchar(30)
,`class_id` tinyint(3) unsigned
,`race_id` tinyint(3) unsigned
,`genre_id` tinyint(3) unsigned
,`head_id` smallint(5) unsigned
,`elo` int(11)
,`level` smallint(5) unsigned
,`exp` int(10) unsigned
,`total_gold` bigint(11) unsigned
,`total_kills` int(6) unsigned
);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `skillpoint`
--

CREATE TABLE `skillpoint` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `value` tinyint(3) UNSIGNED NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `spell`
--

CREATE TABLE `spell` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `spell_id` smallint(5) UNSIGNED DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `statistics`
--

CREATE TABLE `statistics` (
  `name` varchar(50) NOT NULL,
  `value` varchar(50) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `user`
--

CREATE TABLE `user` (
  `id` mediumint(8) UNSIGNED NOT NULL,
  `account_id` mediumint(8) UNSIGNED NOT NULL,
  `deleted` tinyint(1) NOT NULL DEFAULT 0,
  `name` varchar(30) NOT NULL,
  `level` smallint(5) UNSIGNED NOT NULL,
  `exp` int(10) UNSIGNED NOT NULL,
  `elu` int(10) UNSIGNED NOT NULL,
  `genre_id` tinyint(3) UNSIGNED NOT NULL,
  `race_id` tinyint(3) UNSIGNED NOT NULL,
  `class_id` tinyint(3) UNSIGNED NOT NULL,
  `home_id` tinyint(3) UNSIGNED NOT NULL,
  `description` varchar(255) DEFAULT NULL,
  `gold` int(10) UNSIGNED NOT NULL,
  `bank_gold` int(10) UNSIGNED NOT NULL DEFAULT 0,
  `free_skillpoints` smallint(5) UNSIGNED NOT NULL,
  `pets_saved` tinyint(1) NOT NULL DEFAULT 0,
  `votes_amount` smallint(5) UNSIGNED DEFAULT 0,
  `battle_points` smallint(5) UNSIGNED DEFAULT 0,
  `spouse` varchar(30) NOT NULL DEFAULT '',
  `message_info` varchar(512) DEFAULT '',
  `pos_map` smallint(5) UNSIGNED NOT NULL,
  `pos_x` tinyint(3) UNSIGNED NOT NULL,
  `pos_y` tinyint(3) UNSIGNED NOT NULL,
  `last_map` tinyint(3) UNSIGNED NOT NULL DEFAULT 0,
  `body_id` smallint(5) UNSIGNED NOT NULL,
  `head_id` smallint(5) UNSIGNED NOT NULL,
  `weapon_id` smallint(5) UNSIGNED NOT NULL,
  `helmet_id` smallint(5) UNSIGNED NOT NULL,
  `shield_id` smallint(5) UNSIGNED NOT NULL,
  `heading` tinyint(3) UNSIGNED NOT NULL DEFAULT 3,
  `items_amount` tinyint(3) UNSIGNED NOT NULL,
  `slot_armour` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_weapon` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_helmet` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_shield` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_ammo` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_ship` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_mount` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_dm` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_rm` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_magic` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_knuckles` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_tool` tinyint(3) UNSIGNED DEFAULT NULL,
  `min_hp` smallint(5) UNSIGNED NOT NULL,
  `max_hp` smallint(5) UNSIGNED NOT NULL,
  `min_man` smallint(5) UNSIGNED NOT NULL,
  `max_man` smallint(5) UNSIGNED NOT NULL,
  `min_sta` smallint(5) UNSIGNED NOT NULL,
  `max_sta` smallint(5) UNSIGNED NOT NULL,
  `min_ham` smallint(5) UNSIGNED NOT NULL,
  `max_ham` smallint(5) UNSIGNED NOT NULL,
  `min_sed` smallint(5) UNSIGNED NOT NULL,
  `max_sed` smallint(5) UNSIGNED NOT NULL,
  `min_hit` smallint(5) UNSIGNED NOT NULL,
  `max_hit` smallint(5) UNSIGNED NOT NULL,
  `killed_npcs` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `killed_users` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `invent_level` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `is_naked` tinyint(1) NOT NULL DEFAULT 0,
  `is_poisoned` tinyint(1) NOT NULL DEFAULT 0,
  `is_incinerated` tinyint(1) NOT NULL DEFAULT 0,
  `is_hidden` tinyint(1) NOT NULL DEFAULT 0,
  `is_hungry` tinyint(1) NOT NULL DEFAULT 0,
  `is_thirsty` tinyint(1) NOT NULL DEFAULT 0,
  `is_dead` tinyint(1) NOT NULL DEFAULT 0,
  `is_sailing` tinyint(1) NOT NULL DEFAULT 0,
  `is_paralyzed` tinyint(1) NOT NULL DEFAULT 0,
  `is_logged` tinyint(1) NOT NULL DEFAULT 0,
  `is_silenced` tinyint(1) NOT NULL DEFAULT 0,
  `silence_minutes_left` tinyint(4) DEFAULT 0,
  `silence_elapsed_seconds` tinyint(4) DEFAULT 0,
  `is_mounted` tinyint(1) NOT NULL DEFAULT 0,
  `is_banned` tinyint(1) DEFAULT 0,
  `banned_by` varchar(30) NOT NULL DEFAULT '',
  `ban_reason` varchar(255) DEFAULT '',
  `counter_pena` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `deaths` int(10) UNSIGNED NOT NULL DEFAULT 0,
  `pertenece_consejo_real` tinyint(1) NOT NULL DEFAULT 0,
  `pertenece_consejo_caos` tinyint(1) NOT NULL DEFAULT 0,
  `pertenece_real` tinyint(1) NOT NULL DEFAULT 0,
  `pertenece_caos` tinyint(1) NOT NULL DEFAULT 0,
  `ciudadanos_matados` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `criminales_matados` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `recibio_armadura_real` tinyint(1) NOT NULL DEFAULT 0,
  `recibio_armadura_caos` tinyint(1) NOT NULL DEFAULT 0,
  `recibio_exp_real` tinyint(1) NOT NULL DEFAULT 0,
  `recibio_exp_caos` tinyint(1) NOT NULL DEFAULT 0,
  `recompensas_real` tinyint(3) UNSIGNED DEFAULT 0,
  `recompensas_caos` tinyint(3) UNSIGNED DEFAULT 0,
  `reenlistadas` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `fecha_ingreso` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp(),
  `nivel_ingreso` smallint(5) UNSIGNED DEFAULT NULL,
  `matados_ingreso` smallint(5) UNSIGNED DEFAULT NULL,
  `siguiente_recompensa` smallint(5) UNSIGNED DEFAULT NULL,
  `status` tinyint(4) DEFAULT 0,
  `guild_index` smallint(5) UNSIGNED DEFAULT 0,
  `guild_aspirant_index` smallint(5) UNSIGNED DEFAULT NULL,
  `guild_member_history` varchar(1024) DEFAULT NULL,
  `guild_requests_history` varchar(1024) DEFAULT NULL,
  `guild_rejected_because` varchar(255) DEFAULT NULL,
  `chat_global` tinyint(1) DEFAULT 1,
  `chat_combate` tinyint(1) DEFAULT 1,
  `warnings` tinyint(3) UNSIGNED NOT NULL DEFAULT 0,
  `elo` int(11) NOT NULL DEFAULT 1000,
  `return_map` int(11) NOT NULL DEFAULT 0,
  `return_x` int(11) NOT NULL DEFAULT 0,
  `return_y` int(11) NOT NULL DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `whitelist`
--

CREATE TABLE `whitelist` (
  `code` varchar(64) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ROW_FORMAT=DYNAMIC;

-- --------------------------------------------------------

--
-- Estructura para la vista `mao__deal_summaries`
--
DROP TABLE IF EXISTS `mao__deal_summaries`;

CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `mao__deal_summaries`  AS SELECT `d`.`id` AS `id`, `d`.`publisher_account_id` AS `publisher_account_id`, `d`.`character_id` AS `character_id`, `d`.`description` AS `description`, `d`.`deal_type` AS `deal_type`, `d`.`gold_price` AS `gold_price`, `d`.`closed_at` AS `closed_at`, `d`.`created_at` AS `created_at`, `d`.`updated_at` AS `updated_at`, `u`.`name` AS `character_name`, `u`.`class_id` AS `class_id`, `u`.`race_id` AS `race_id`, `u`.`genre_id` AS `genre_id`, `u`.`elo` AS `elo`, `u`.`level` AS `level`, `u`.`head_id` AS `head_id`, `u`.`max_hp` AS `max_hp`, `u`.`warnings` AS `warnings` FROM (`mao__deals` `d` join `user` `u` on(`d`.`character_id` = `u`.`id`)) ;

-- --------------------------------------------------------

--
-- Estructura para la vista `mao__offer_summaries`
--
DROP TABLE IF EXISTS `mao__offer_summaries`;

CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `mao__offer_summaries`  AS SELECT `o`.`id` AS `id`, `o`.`deal_id` AS `deal_id`, `o`.`offeror_account_id` AS `offeror_account_id`, `o`.`character_id` AS `character_id`, `o`.`offer_price` AS `offer_price`, `o`.`state` AS `state`, `o`.`description` AS `description`, `o`.`created_at` AS `created_at`, `o`.`updated_at` AS `updated_at`, `u`.`name` AS `character_name`, `u`.`class_id` AS `class_id`, `u`.`race_id` AS `race_id`, `u`.`genre_id` AS `genre_id`, `u`.`elo` AS `elo`, `u`.`level` AS `level`, `u`.`head_id` AS `head_id`, `u`.`max_hp` AS `max_hp`, `u`.`warnings` AS `warnings` FROM (`mao__offers` `o` join `user` `u` on(`o`.`character_id` = `u`.`id`)) ;

-- --------------------------------------------------------

--
-- Estructura para la vista `ranking_users`
--
DROP TABLE IF EXISTS `ranking_users`;

CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `ranking_users`  AS SELECT `u`.`id` AS `id`, `u`.`name` AS `character_name`, `u`.`class_id` AS `class_id`, `u`.`race_id` AS `race_id`, `u`.`genre_id` AS `genre_id`, `u`.`head_id` AS `head_id`, `u`.`elo` AS `elo`, `u`.`level` AS `level`, `u`.`exp` AS `exp`, `u`.`gold`+ `u`.`bank_gold` AS `total_gold`, `u`.`criminales_matados`+ `u`.`ciudadanos_matados` AS `total_kills` FROM `user` AS `u` WHERE `u`.`deleted` = 0 ;

--
-- Índices para tablas volcadas
--

--
-- Indices de la tabla `account`
--
ALTER TABLE `account`
  ADD PRIMARY KEY (`id`),
  ADD KEY `email` (`email`);

--
-- Indices de la tabla `attribute`
--
ALTER TABLE `attribute`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `bank_item`
--
ALTER TABLE `bank_item`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `connection`
--
ALTER TABLE `connection`
  ADD PRIMARY KEY (`user_id`,`ip`);

--
-- Indices de la tabla `house_key`
--
ALTER TABLE `house_key`
  ADD PRIMARY KEY (`key_obj`),
  ADD KEY `fk_account` (`account_id`);

--
-- Indices de la tabla `inventory_item`
--
ALTER TABLE `inventory_item`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `mail`
--
ALTER TABLE `mail`
  ADD PRIMARY KEY (`id`),
  ADD KEY `user_id` (`user_id`);

--
-- Indices de la tabla `mao__claims`
--
ALTER TABLE `mao__claims`
  ADD PRIMARY KEY (`id`),
  ADD KEY `FK_mao__claims_deal_id` (`deal_id`),
  ADD KEY `FK_mao__claims_account_id` (`claimer_account_id`),
  ADD KEY `FK_mao__claims_user_id` (`claimer_character_id`);

--
-- Indices de la tabla `mao__deals`
--
ALTER TABLE `mao__deals`
  ADD PRIMARY KEY (`id`),
  ADD KEY `FK_account_id` (`publisher_account_id`),
  ADD KEY `FK_user_id` (`character_id`);

--
-- Indices de la tabla `mao__offers`
--
ALTER TABLE `mao__offers`
  ADD PRIMARY KEY (`id`),
  ADD KEY `FK_mao__offers_deal_id` (`deal_id`),
  ADD KEY `FK_mao__offers_account_id` (`offeror_account_id`),
  ADD KEY `FK_mao__offers_user_id` (`character_id`);

--
-- Indices de la tabla `pet`
--
ALTER TABLE `pet`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `punishment`
--
ALTER TABLE `punishment`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `quest`
--
ALTER TABLE `quest`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `quest_done`
--
ALTER TABLE `quest_done`
  ADD PRIMARY KEY (`user_id`,`quest_id`);

--
-- Indices de la tabla `skillpoint`
--
ALTER TABLE `skillpoint`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `spell`
--
ALTER TABLE `spell`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `statistics`
--
ALTER TABLE `statistics`
  ADD PRIMARY KEY (`name`);

--
-- Indices de la tabla `user`
--
ALTER TABLE `user`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `index_table_user` (`id`,`account_id`,`deleted`),
  ADD KEY `fk_user_account` (`account_id`),
  ADD KEY `name` (`name`);

--
-- Indices de la tabla `whitelist`
--
ALTER TABLE `whitelist`
  ADD PRIMARY KEY (`code`) USING BTREE;

--
-- AUTO_INCREMENT de las tablas volcadas
--

--
-- AUTO_INCREMENT de la tabla `account`
--
ALTER TABLE `account`
  MODIFY `id` mediumint(8) UNSIGNED NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT de la tabla `mail`
--
ALTER TABLE `mail`
  MODIFY `id` mediumint(8) UNSIGNED NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT de la tabla `mao__claims`
--
ALTER TABLE `mao__claims`
  MODIFY `id` int(11) UNSIGNED NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT de la tabla `mao__deals`
--
ALTER TABLE `mao__deals`
  MODIFY `id` int(11) UNSIGNED NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT de la tabla `mao__offers`
--
ALTER TABLE `mao__offers`
  MODIFY `id` int(11) UNSIGNED NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT de la tabla `user`
--
ALTER TABLE `user`
  MODIFY `id` mediumint(8) UNSIGNED NOT NULL AUTO_INCREMENT;

--
-- Restricciones para tablas volcadas
--

--
-- Filtros para la tabla `attribute`
--
ALTER TABLE `attribute`
  ADD CONSTRAINT `fk_attribute_user` FOREIGN KEY (`user_id`) REFERENCES `user` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Filtros para la tabla `bank_item`
--
ALTER TABLE `bank_item`
  ADD CONSTRAINT `fk_bank_user` FOREIGN KEY (`user_id`) REFERENCES `user` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Filtros para la tabla `connection`
--
ALTER TABLE `connection`
  ADD CONSTRAINT `fk_ip_user` FOREIGN KEY (`user_id`) REFERENCES `user` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Filtros para la tabla `house_key`
--
ALTER TABLE `house_key`
  ADD CONSTRAINT `fk_account` FOREIGN KEY (`account_id`) REFERENCES `account` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Filtros para la tabla `inventory_item`
--
ALTER TABLE `inventory_item`
  ADD CONSTRAINT `fk_inventory_user` FOREIGN KEY (`user_id`) REFERENCES `user` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Filtros para la tabla `mail`
--
ALTER TABLE `mail`
  ADD CONSTRAINT `fk_mail` FOREIGN KEY (`user_id`) REFERENCES `user` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Filtros para la tabla `mao__claims`
--
ALTER TABLE `mao__claims`
  ADD CONSTRAINT `FK_mao__claims_account_id` FOREIGN KEY (`claimer_account_id`) REFERENCES `account` (`id`),
  ADD CONSTRAINT `FK_mao__claims_deal_id` FOREIGN KEY (`deal_id`) REFERENCES `mao__deals` (`id`),
  ADD CONSTRAINT `FK_mao__claims_user_id` FOREIGN KEY (`claimer_character_id`) REFERENCES `user` (`id`);

--
-- Filtros para la tabla `mao__deals`
--
ALTER TABLE `mao__deals`
  ADD CONSTRAINT `FK_account_id` FOREIGN KEY (`publisher_account_id`) REFERENCES `account` (`id`),
  ADD CONSTRAINT `FK_user_id` FOREIGN KEY (`character_id`) REFERENCES `user` (`id`);

--
-- Filtros para la tabla `mao__offers`
--
ALTER TABLE `mao__offers`
  ADD CONSTRAINT `FK_mao__offers_account_id` FOREIGN KEY (`offeror_account_id`) REFERENCES `account` (`id`),
  ADD CONSTRAINT `FK_mao__offers_deal_id` FOREIGN KEY (`deal_id`) REFERENCES `mao__deals` (`id`),
  ADD CONSTRAINT `FK_mao__offers_user_id` FOREIGN KEY (`character_id`) REFERENCES `user` (`id`);

--
-- Filtros para la tabla `pet`
--
ALTER TABLE `pet`
  ADD CONSTRAINT `fk_pet_user` FOREIGN KEY (`user_id`) REFERENCES `user` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Filtros para la tabla `punishment`
--
ALTER TABLE `punishment`
  ADD CONSTRAINT `fk_punishment_user` FOREIGN KEY (`user_id`) REFERENCES `user` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Filtros para la tabla `quest`
--
ALTER TABLE `quest`
  ADD CONSTRAINT `fk_quest_user` FOREIGN KEY (`user_id`) REFERENCES `user` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Filtros para la tabla `quest_done`
--
ALTER TABLE `quest_done`
  ADD CONSTRAINT `fk_quest_done_user` FOREIGN KEY (`user_id`) REFERENCES `user` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Filtros para la tabla `skillpoint`
--
ALTER TABLE `skillpoint`
  ADD CONSTRAINT `fk_skillpoint_user` FOREIGN KEY (`user_id`) REFERENCES `user` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Filtros para la tabla `spell`
--
ALTER TABLE `spell`
  ADD CONSTRAINT `fk_spell_user` FOREIGN KEY (`user_id`) REFERENCES `user` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Filtros para la tabla `user`
--
ALTER TABLE `user`
  ADD CONSTRAINT `fk_user_account` FOREIGN KEY (`account_id`) REFERENCES `account` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
