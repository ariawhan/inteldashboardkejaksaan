CREATE TABLE IF NOT EXISTS backdated_number_reservations (
    id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    reservation_token CHAR(32) NOT NULL UNIQUE,
    document_type VARCHAR(30) NOT NULL,
    report_id INT UNSIGNED NOT NULL,
    report_date DATE NOT NULL,
    sequence_label VARCHAR(30) NOT NULL,
    created_by INT UNSIGNED NOT NULL,
    reserved_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE KEY uq_backdated_number (document_type, report_date, sequence_label),
    UNIQUE KEY uq_backdated_report (document_type, report_id),
    CONSTRAINT fk_backdated_number_user FOREIGN KEY (created_by)
        REFERENCES users(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
