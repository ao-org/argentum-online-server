DELETE FROM pet WHERE user_id IN (
    SELECT p.user_id
    FROM pet p
    LEFT JOIN "user" u ON u.id = p.user_id
    WHERE u.id IS NULL
);
