-- 1. Borrar miembros que no son líderes
DELETE FROM guild_members
WHERE user_id NOT IN (
  SELECT leader_id
  FROM guilds
  WHERE leader_id IS NOT NULL
);