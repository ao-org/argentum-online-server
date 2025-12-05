-- 1. Resetear guild_index de los usuarios que no son líderes
UPDATE user
SET guild_index = 0
WHERE id NOT IN (
  SELECT leader_id
  FROM guilds
  WHERE leader_id IS NOT NULL
);