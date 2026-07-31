ALTER TABLE lapinhar_reports
MODIFY COLUMN sipede_status ENUM('belum', 'sudah', 'manual') NOT NULL DEFAULT 'belum';

ALTER TABLE lapinsus_reports
MODIFY COLUMN sipede_status ENUM('belum', 'sudah', 'manual') NOT NULL DEFAULT 'belum';

DROP VIEW IF EXISTS reports;

CREATE VIEW reports AS
SELECT * FROM lapinhar_reports
UNION ALL
SELECT * FROM lapinsus_reports;
