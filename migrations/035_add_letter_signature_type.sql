ALTER TABLE lapinhar_reports
ADD COLUMN IF NOT EXISTS letter_signature_type VARCHAR(20) NOT NULL DEFAULT 'tte';

ALTER TABLE lapinsus_reports
ADD COLUMN IF NOT EXISTS letter_signature_type VARCHAR(20) NOT NULL DEFAULT 'tte';

DROP VIEW IF EXISTS reports;
CREATE VIEW reports AS
SELECT * FROM lapinhar_reports
UNION ALL
SELECT * FROM lapinsus_reports;
