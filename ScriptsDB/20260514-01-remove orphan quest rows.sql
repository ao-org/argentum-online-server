DELETE FROM quest WHERE user_id IN (
    SELECT q.user_id
    FROM quest q
    LEFT JOIN "user" u ON u.id = q.user_id
    WHERE u.id IS NULL
);
