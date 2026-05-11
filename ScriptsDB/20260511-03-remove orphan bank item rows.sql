DELETE FROM bank_item WHERE user_id IN (
    SELECT b.user_id
    FROM bank_item b
    LEFT JOIN "user" u ON u.id = b.user_id
    WHERE u.id IS NULL
);
