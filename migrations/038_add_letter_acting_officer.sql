ALTER TABLE lapinhar_reports
ADD COLUMN IF NOT EXISTS acting_officer_type VARCHAR(10) NULL,
ADD COLUMN IF NOT EXISTS acting_officer_name VARCHAR(200) NULL,
ADD COLUMN IF NOT EXISTS acting_officer_position VARCHAR(200) NULL,
ADD COLUMN IF NOT EXISTS acting_officer_nip VARCHAR(255) NULL;

ALTER TABLE lapinsus_reports
ADD COLUMN IF NOT EXISTS acting_officer_type VARCHAR(10) NULL,
ADD COLUMN IF NOT EXISTS acting_officer_name VARCHAR(200) NULL,
ADD COLUMN IF NOT EXISTS acting_officer_position VARCHAR(200) NULL,
ADD COLUMN IF NOT EXISTS acting_officer_nip VARCHAR(255) NULL;

DROP VIEW IF EXISTS reports;
CREATE VIEW reports AS
SELECT * FROM lapinhar_reports
UNION ALL
SELECT * FROM lapinsus_reports;
