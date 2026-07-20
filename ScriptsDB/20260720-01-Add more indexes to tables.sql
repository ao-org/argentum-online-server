CREATE INDEX idx_pet_user_id ON pet(user_id);
CREATE INDEX idx_spell_user_id ON spell(user_id);
CREATE INDEX idx_skillpoint_user_id ON skillpoint(user_id);
CREATE INDEX idx_inventory_item_user_id ON inventory_item(user_id);
CREATE INDEX idx_bank_item_user_id ON bank_item(user_id);
CREATE INDEX idx_accounts ON account(id);