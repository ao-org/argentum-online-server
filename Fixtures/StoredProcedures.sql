-- --------------------------------------------------------
-- Host:                         ao20-testeo-tercero.duckdns.org
-- Server version:               10.4.14-MariaDB-log - mariadb.org binary distribution
-- Server OS:                    Win64
-- HeidiSQL Version:             11.3.0.6295
-- --------------------------------------------------------

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET NAMES utf8 */;
/*!50503 SET NAMES utf8mb4 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

-- Dumping structure for procedure ao_server_test.sp_LoadChar
DELIMITER //
CREATE PROCEDURE `sp_LoadChar`(IN Username VARCHAR(30), IN Connect BOOL)
Main:BEGIN
   DECLARE IsBanned TINYINT(1);
	DECLARE UID TYPE OF user.id;
	DECLARE AID TYPE OF user.account_id;
	
	SELECT 
		@UID := UINFO.id id,
		@AID := UINFO.account_id account_id, 
		UINFO.name,
		UINFO.level,
		UINFO.exp,
		UINFO.genre_id,
		UINFO.race_id,
		UINFO.class_id,
		UINFO.home_id,
		UINFO.description,
		UINFO.gold,
		UINFO.bank_gold,
		UINFO.free_skillpoints,
		UINFO.pos_map,
		UINFO.pos_x,
		UINFO.pos_y,
		UINFO.message_info,
		UINFO.body_id,
		UINFO.head_id,
		UINFO.weapon_id,
		UINFO.helmet_id,
		UINFO.shield_id,
		UINFO.heading,
		UINFO.slot_armour,
		UINFO.slot_weapon,
		UINFO.slot_helmet,
		UINFO.slot_shield,
		UINFO.slot_ammo,
		UINFO.slot_ship,
		UINFO.slot_mount,
		UINFO.slot_dm,
		UINFO.slot_rm,
		UINFO.slot_knuckles,
		UINFO.slot_tool,
		UINFO.slot_magic,
		UINFO.min_hp,
		UINFO.max_hp,
		UINFO.min_man,
		UINFO.max_man,
		UINFO.min_sta,
		UINFO.max_sta,
		UINFO.min_ham,
		UINFO.max_ham,
		UINFO.min_sed,
		UINFO.max_sed,
		UINFO.min_hit,
		UINFO.max_hit,
		UINFO.killed_npcs,
		UINFO.killed_users,
		UINFO.invent_level,
		UINFO.ELO,
		UINFO.is_naked,
		UINFO.is_poisoned,
		UINFO.is_incinerated,
		@IsBanned := UINFO.is_banned is_banned,
		UINFO.ban_reason,
		UINFO.banned_by,
		UINFO.is_dead,
		UINFO.is_sailing,
		UINFO.is_paralyzed,
		UINFO.deaths,
		UINFO.is_mounted,
		UINFO.spouse,
		UINFO.is_silenced,
		UINFO.silence_minutes_left,
		UINFO.silence_elapsed_seconds,
		UINFO.pets_saved,
		UINFO.return_map,
		UINFO.return_x,
		UINFO.return_y,
		UINFO.counter_pena,
		UINFO.chat_global,
		UINFO.chat_combate,
		UINFO.pertenece_consejo_real,
		UINFO.pertenece_consejo_caos,
		UINFO.pertenece_real,
		UINFO.pertenece_caos,
		UINFO.ciudadanos_matados,
		UINFO.criminales_matados,
		UINFO.recibio_armadura_real,
		UINFO.recibio_armadura_caos,
		UINFO.recibio_exp_real,
		UINFO.recibio_exp_caos,
		UINFO.recompensas_real,
		UINFO.recompensas_caos,
		UINFO.reenlistadas,
		UINFO.nivel_ingreso,
		UINFO.matados_ingreso,
		UINFO.siguiente_recompensa,
		UINFO.status,
		UINFO.guild_index,
		UINFO.guild_rejected_because,
		UINFO.warnings,
		DATE_FORMAT(UINFO.fecha_ingreso, '%Y-%m-%d') as 'fecha_ingreso_format',
		UATTR.strength,
		UATTR.agility,
		UATTR.constitution,
		UATTR.intelligence,
		UATTR.charisma
	FROM
		user UINFO
			INNER JOIN attribute UATTR
				ON UINFO.id = UATTR.user_id
	WHERE
		UINFO.name = Username;
	
	IF @IsBanned > 0 THEN
		LEAVE Main;
	END IF;
	
	-- User spells
	SELECT number, spell_id FROM spell WHERE user_id = @UID;
	
	-- User pets
	SELECT number, pet_id FROM pet WHERE user_id = @UID;
	
	-- User inventory
	SELECT number, item_id, is_equipped, amount FROM inventory_item WHERE user_id = @UID;
	
	-- User bank 
	SELECT number, item_id, amount FROM bank_item WHERE user_id = @UID;
		
	-- User skills 
	SELECT number, value FROM skillpoint WHERE user_id = @UID;
			
	-- User quests 
	SELECT number, quest_id, npcs, npcstarget FROM quest WHERE user_id = @UID;
			
	-- User quests done 
	SELECT quest_id FROM quest_done WHERE user_id = @UID;

	-- Keys
	SELECT key_obj FROM house_key WHERE account_id = @AID;
	
	-- User's connected status
	UPDATE user SET is_logged = true WHERE id = @UID;
END//
DELIMITER ;

/*!40101 SET SQL_MODE=IFNULL(@OLD_SQL_MODE, '') */;
/*!40014 SET FOREIGN_KEY_CHECKS=IFNULL(@OLD_FOREIGN_KEY_CHECKS, 1) */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40111 SET SQL_NOTES=IFNULL(@OLD_SQL_NOTES, 1) */;
