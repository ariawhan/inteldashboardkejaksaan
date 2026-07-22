DROP VIEW IF EXISTS reports;

CREATE VIEW reports AS
SELECT * FROM lapinhar_reports
UNION ALL
SELECT * FROM lapinsus_reports;
