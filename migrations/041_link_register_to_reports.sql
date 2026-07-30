ALTER TABLE register_intelijen_entries
ADD COLUMN IF NOT EXISTS source_report_type VARCHAR(20) NULL AFTER register_code,
ADD COLUMN IF NOT EXISTS source_report_id INT NULL AFTER source_report_type;

ALTER TABLE register_intelijen_entries
ADD UNIQUE INDEX uq_register_intelijen_source_report
(source_report_type, source_report_id);
