ALTER TABLE lapinhar_reports
ADD COLUMN IF NOT EXISTS use_scanned_signatures TINYINT(1) NOT NULL DEFAULT 1;

ALTER TABLE lapinsus_reports
ADD COLUMN IF NOT EXISTS use_scanned_signatures TINYINT(1) NOT NULL DEFAULT 1;

DROP VIEW IF EXISTS reports;
CREATE VIEW reports AS
SELECT * FROM lapinhar_reports
UNION ALL
SELECT * FROM lapinsus_reports;
