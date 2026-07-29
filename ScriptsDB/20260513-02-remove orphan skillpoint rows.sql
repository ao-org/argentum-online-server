DELETE FROM skillpoint WHERE user_id IN (
    SELECT s.user_id
    FROM skillpoint s
    LEFT JOIN "user" u ON u.id = s.user_id
    WHERE u.id IS NULL
);
