DELETE FROM quest_done WHERE user_id IN (
    SELECT qd.user_id
    FROM quest_done qd
    LEFT JOIN "user" u ON u.id = qd.user_id
    WHERE u.id IS NULL
);
