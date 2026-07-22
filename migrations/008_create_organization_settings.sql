CREATE TABLE IF NOT EXISTS organization_settings (
    id TINYINT UNSIGNED PRIMARY KEY,
    organization_name VARCHAR(255) NOT NULL,
    address TEXT NOT NULL,
    phone VARCHAR(100) NOT NULL DEFAULT '',
    website VARCHAR(255) NOT NULL DEFAULT '',
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

INSERT IGNORE INTO organization_settings (id, organization_name, address, phone, website) VALUES
(1, 'KEJAKSAAN NEGERI BULELENG', '', '', '');
