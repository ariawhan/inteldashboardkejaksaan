ALTER TABLE document_counters
ADD COLUMN IF NOT EXISTS document_year SMALLINT UNSIGNED NOT NULL DEFAULT 2026 AFTER document_type;

ALTER TABLE document_counters DROP PRIMARY KEY;
ALTER TABLE document_counters ADD PRIMARY KEY (document_type, document_year);

ALTER TABLE document_number_reservations
ADD COLUMN IF NOT EXISTS document_year SMALLINT UNSIGNED NOT NULL DEFAULT 2026 AFTER document_type;

ALTER TABLE document_number_reservations DROP INDEX uq_document_sequence;
ALTER TABLE document_number_reservations
ADD UNIQUE INDEX uq_document_year_sequence (document_type, document_year, sequence_number);
