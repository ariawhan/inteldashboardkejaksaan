ALTER TABLE signatories
MODIFY COLUMN position_code ENUM('kajari', 'kasi_intel', 'kasubsi_1', 'kasubsi_2') NOT NULL UNIQUE;

INSERT IGNORE INTO signatories (position_code, position_name, full_name, rank_nip)
VALUES (
    'kajari',
    'Kepala Kejaksaan Negeri Buleleng',
    'DICKY DARMAWAN, S.H., M.H.',
    'Jaksa Utama Pratama NIP. 19711201 199903 1 009'
);
