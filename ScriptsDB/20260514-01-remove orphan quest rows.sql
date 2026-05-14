DELETE FROM quest WHERE user_id IN (
    SELECT q.user_id
    FROM quest q
    LEFT JOIN "user" u ON u.id = q.user_id
    WHERE u.id IS NULL
);

DELETE FROM quest_done WHERE user_id IN (
    SELECT qd.user_id
    FROM quest_done qd
    LEFT JOIN "user" u ON u.id = qd.user_id
    WHERE u.id IS NULL
);
