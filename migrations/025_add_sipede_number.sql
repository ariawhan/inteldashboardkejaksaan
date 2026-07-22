ALTER TABLE lapinhar_reports
ADD COLUMN IF NOT EXISTS sipede_number VARCHAR(150) NULL AFTER sipede_status;

ALTER TABLE lapinsus_reports
ADD COLUMN IF NOT EXISTS sipede_number VARCHAR(150) NULL AFTER sipede_status;

DROP VIEW IF EXISTS reports;

CREATE VIEW reports AS
SELECT * FROM lapinhar_reports
UNION ALL
SELECT * FROM lapinsus_reports;
