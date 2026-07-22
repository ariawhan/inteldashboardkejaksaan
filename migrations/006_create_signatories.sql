CREATE TABLE IF NOT EXISTS signatories (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    position_code ENUM('kasi_intel', 'kasubsi_1', 'kasubsi_2') NOT NULL UNIQUE,
    position_name VARCHAR(150) NOT NULL,
    full_name VARCHAR(200) NOT NULL DEFAULT '',
    rank_nip VARCHAR(255) NOT NULL DEFAULT '',
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

INSERT IGNORE INTO signatories (position_code, position_name, full_name, rank_nip) VALUES
('kasi_intel', 'Kepala Seksi Intelijen', 'I Dewa Gede Baskara Haryasa, S.H.', 'Jaksa Muda NIP. 19761015 200312 1 006'),
('kasubsi_1', 'Kepala Subseksi I Intelijen', '', ''),
('kasubsi_2', 'Kepala Subseksi II Intelijen', 'I Nyoman Arya Wira Temaja, S.H.', 'Ajun Jaksa NIP. 19960726 201902 1 001');
