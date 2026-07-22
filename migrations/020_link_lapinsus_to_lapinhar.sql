ALTER TABLE lapinhar_reports
ADD COLUMN IF NOT EXISTS source_lapinhar_id INT UNSIGNED NULL AFTER created_by;

ALTER TABLE lapinsus_reports
ADD COLUMN IF NOT EXISTS source_lapinhar_id INT UNSIGNED NULL AFTER created_by;

ALTER TABLE lapinsus_reports
ADD UNIQUE INDEX IF NOT EXISTS uq_lapinsus_source_lapinhar (source_lapinhar_id);
