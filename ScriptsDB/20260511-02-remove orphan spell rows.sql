DELETE FROM spell WHERE user_id IN (SELECT s.user_id FROM spell s LEFT JOIN "user" u ON u.id = s.user_id WHERE u.id IS NULL);
