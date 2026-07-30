ALTER TABLE lapinhar_reports
ADD COLUMN IF NOT EXISTS letter_use_digital_stamp TINYINT(1) NOT NULL DEFAULT 0;

ALTER TABLE lapinsus_reports
ADD COLUMN IF NOT EXISTS letter_use_digital_stamp TINYINT(1) NOT NULL DEFAULT 0;

DROP VIEW IF EXISTS reports;
CREATE VIEW reports AS
SELECT * FROM lapinhar_reports
UNION ALL
SELECT * FROM lapinsus_reports;
