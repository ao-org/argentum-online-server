DELETE FROM inventory_item WHERE user_id IN (
    SELECT i.user_id
    FROM inventory_item i
    LEFT JOIN "user" u ON u.id = i.user_id
    WHERE u.id IS NULL
);
