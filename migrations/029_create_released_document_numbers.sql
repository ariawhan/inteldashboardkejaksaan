CREATE TABLE IF NOT EXISTS released_document_numbers (
    id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    document_type VARCHAR(30) NOT NULL,
    document_year SMALLINT UNSIGNED NOT NULL,
    sequence_number INT UNSIGNED NOT NULL,
    released_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE KEY uq_released_document_number (document_type, document_year, sequence_number)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
