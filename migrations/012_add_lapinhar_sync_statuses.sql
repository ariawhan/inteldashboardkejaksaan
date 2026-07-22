ALTER TABLE reports
    ADD COLUMN IF NOT EXISTS inteliz_status ENUM('belum', 'sudah') NOT NULL DEFAULT 'belum' AFTER status,
    ADD COLUMN IF NOT EXISTS lapinsus_status ENUM('belum', 'sudah') NOT NULL DEFAULT 'belum' AFTER inteliz_status;
