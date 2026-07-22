ALTER TABLE inteliz_user_settings
    ADD COLUMN IF NOT EXISTS session_data_encrypted LONGTEXT NULL AFTER inteliz_password_encrypted,
    ADD COLUMN IF NOT EXISTS connected_at TIMESTAMP NULL AFTER session_data_encrypted;
