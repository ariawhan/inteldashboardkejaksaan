ALTER TABLE signatories
ADD COLUMN IF NOT EXISTS signature_image VARCHAR(255) NULL AFTER use_tte;
