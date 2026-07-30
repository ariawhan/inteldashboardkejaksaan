ALTER TABLE lapinhar_reports
ADD COLUMN IF NOT EXISTS acting_officer_rank VARCHAR(150) NULL;

ALTER TABLE lapinsus_reports
ADD COLUMN IF NOT EXISTS acting_officer_rank VARCHAR(150) NULL;

DROP VIEW IF EXISTS reports;
CREATE VIEW reports AS
SELECT * FROM lapinhar_reports
UNION ALL
SELECT * FROM lapinsus_reports;
