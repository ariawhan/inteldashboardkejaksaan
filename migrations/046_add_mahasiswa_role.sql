ALTER TABLE users
    MODIFY COLUMN role ENUM('admin', 'user', 'mahasiswa') NOT NULL DEFAULT 'user';
