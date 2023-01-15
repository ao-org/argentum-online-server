CREATE view `ranking_users`
AS
  SELECT `u`.`id`                     AS `id`,
         `u`.`name`                   AS `character_name`,
         `u`.`class_id`               AS `class_id`,
         `u`.`race_id`                AS `race_id`,
         `u`.`genre_id`               AS `genre_id`,
         `u`.`head_id`                AS `head_id`,
         `u`.`level`                  AS `level`,
         `u`.`exp`                    AS `exp`,
         `u`.`gold` + `u`.`bank_gold` AS `total_gold`,
         `u`.`criminales_matados` + `u`.`ciudadanos_matados`   AS `total_kills`,
         `u`.`criminales_matados`     AS `criminales_matados`,
         `u`.`ciudadanos_matados`     AS `ciudadanos_matados`,
         `u`.`puntos_pesca`           AS `puntos_pesca`,
         `u`.`deaths`                 AS `deaths`,
         `u`.`killed_npcs`            AS `killed_npcs`
  FROM   `user` `u`
  WHERE  `u`.`deleted` <> true AND `u`.`is_banned` <> true
