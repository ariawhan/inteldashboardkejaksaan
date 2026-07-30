ALTER TABLE lapinhar_reports
ADD COLUMN IF NOT EXISTS report_acting_type VARCHAR(10) NULL AFTER use_digital_stamp,
ADD COLUMN IF NOT EXISTS report_acting_name VARCHAR(200) NULL AFTER report_acting_type,
ADD COLUMN IF NOT EXISTS report_acting_position VARCHAR(200) NULL AFTER report_acting_name,
ADD COLUMN IF NOT EXISTS report_acting_nip VARCHAR(255) NULL AFTER report_acting_position;

ALTER TABLE lapinsus_reports
ADD COLUMN IF NOT EXISTS report_acting_type VARCHAR(10) NULL AFTER use_digital_stamp,
ADD COLUMN IF NOT EXISTS report_acting_name VARCHAR(200) NULL AFTER report_acting_type,
ADD COLUMN IF NOT EXISTS report_acting_position VARCHAR(200) NULL AFTER report_acting_name,
ADD COLUMN IF NOT EXISTS report_acting_nip VARCHAR(255) NULL AFTER report_acting_position;
