CREATE TABLE IF NOT EXISTS register_intelijen_generic_entries (
    id INT AUTO_INCREMENT PRIMARY KEY,
    register_code VARCHAR(20) NOT NULL,
    entry_date DATE NULL,
    chart_label VARCHAR(255) NULL,
    payload LONGTEXT NOT NULL,
    created_by INT UNSIGNED NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    CONSTRAINT fk_register_intelijen_generic_created_by FOREIGN KEY (created_by) REFERENCES users(id) ON DELETE CASCADE,
    INDEX idx_register_intelijen_generic_code_date (register_code, entry_date),
    INDEX idx_register_intelijen_generic_chart (register_code, chart_label),
    INDEX idx_register_intelijen_generic_created_by (created_by)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
