ALTER TABLE sipede_user_settings
ADD COLUMN IF NOT EXISTS session_data_encrypted LONGTEXT NULL AFTER sipede_password_encrypted;
