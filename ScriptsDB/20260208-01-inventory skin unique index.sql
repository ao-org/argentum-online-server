BEGIN TRANSACTION;

DELETE FROM inventory_item_skins
WHERE id NOT IN (
    SELECT MAX(id)
    FROM inventory_item_skins
    GROUP BY user_id, skin_id
);

CREATE UNIQUE INDEX IF NOT EXISTS uq_inventory_item_skins_user_skin
ON inventory_item_skins(user_id, skin_id);

COMMIT;
