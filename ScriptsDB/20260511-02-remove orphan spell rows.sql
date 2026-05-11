DELETE FROM spell
WHERE NOT EXISTS (
    SELECT 1
    FROM "user" u
    WHERE u.id = spell.user_id
);
