ALTER TABLE organization_settings
ADD COLUMN IF NOT EXISTS digital_stamp VARCHAR(255) NULL AFTER website;
