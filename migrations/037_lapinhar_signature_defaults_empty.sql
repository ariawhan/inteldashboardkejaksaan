ALTER TABLE lapinhar_reports
ALTER COLUMN use_scanned_signatures SET DEFAULT 0;

ALTER TABLE lapinhar_reports
ALTER COLUMN use_digital_stamp SET DEFAULT 0;

UPDATE lapinhar_reports
SET use_scanned_signatures=0, use_digital_stamp=0
WHERE status='draft';
