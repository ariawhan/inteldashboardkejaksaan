ALTER TABLE lapinhar_reports
ADD COLUMN IF NOT EXISTS sipede_status ENUM('belum', 'sudah') NOT NULL DEFAULT 'belum' AFTER inteliz_status;

ALTER TABLE lapinsus_reports
ADD COLUMN IF NOT EXISTS sipede_status ENUM('belum', 'sudah') NOT NULL DEFAULT 'belum' AFTER inteliz_status;
