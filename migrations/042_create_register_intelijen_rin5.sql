CREATE TABLE IF NOT EXISTS register_intelijen_rin5_entries (
    id INT AUTO_INCREMENT PRIMARY KEY,
    intelligence_product_type VARCHAR(255) NOT NULL,
    intelligence_product_number VARCHAR(255) NOT NULL,
    intelligence_product_date DATE NOT NULL,
    field_code VARCHAR(20) NOT NULL,
    subject TEXT NOT NULL,
    leader_disposition TEXT NULL,
    remarks VARCHAR(255) NULL,
    created_by INT UNSIGNED NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    CONSTRAINT fk_register_intelijen_rin5_created_by FOREIGN KEY (created_by) REFERENCES users(id) ON DELETE CASCADE,
    INDEX idx_register_intelijen_rin5_product_date (intelligence_product_date),
    INDEX idx_register_intelijen_rin5_field_code (field_code),
    INDEX idx_register_intelijen_rin5_created_by (created_by)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
