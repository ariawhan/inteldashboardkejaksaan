ALTER TABLE register_intelijen_rin5_entries
ADD COLUMN IF NOT EXISTS source_report_type VARCHAR(20) NULL AFTER id,
ADD COLUMN IF NOT EXISTS source_report_id INT NULL AFTER source_report_type;

ALTER TABLE register_intelijen_rin5_entries
ADD UNIQUE INDEX IF NOT EXISTS uq_register_rin5_source_report
(source_report_type, source_report_id);
