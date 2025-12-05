-- 1. Borrar miembros que no son líderes
DELETE FROM guild_members
WHERE user_id NOT IN (
  SELECT leader_id
  FROM guilds
  WHERE leader_id IS NOT NULL
);

-- 2. Resetear guild_index de los usuarios que no son líderes
UPDATE user
SET guild_index = 0
WHERE id NOT IN (
  SELECT leader_id
  FROM guilds
  WHERE leader_id IS NOT NULL
);