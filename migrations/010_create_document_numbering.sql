ALTER TABLE organization_settings ADD COLUMN IF NOT EXISTS institution_code VARCHAR(50) NOT NULL DEFAULT 'N.1.11' AFTER organization_name;

CREATE TABLE IF NOT EXISTS document_counters (
    document_type VARCHAR(30) PRIMARY KEY,
    next_number INT UNSIGNED NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

INSERT IGNORE INTO document_counters (document_type, next_number) VALUES ('lapinhar', 229);

CREATE TABLE IF NOT EXISTS document_number_reservations (
    id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    reservation_token CHAR(32) NOT NULL UNIQUE,
    document_type VARCHAR(30) NOT NULL,
    sequence_number INT UNSIGNED NOT NULL,
    created_by INT UNSIGNED NOT NULL,
    status ENUM('reserved', 'used') NOT NULL DEFAULT 'reserved',
    report_id INT UNSIGNED NULL,
    reserved_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    used_at TIMESTAMP NULL,
    UNIQUE KEY uq_document_sequence (document_type, sequence_number),
    CONSTRAINT fk_number_reservation_user FOREIGN KEY (created_by) REFERENCES users(id) ON DELETE RESTRICT,
    CONSTRAINT fk_number_reservation_report FOREIGN KEY (report_id) REFERENCES reports(id) ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

ALTER TABLE reports ADD UNIQUE INDEX IF NOT EXISTS uq_reports_report_number (report_number);
