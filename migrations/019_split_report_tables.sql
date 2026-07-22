CREATE TABLE IF NOT EXISTS lapinhar_reports LIKE reports;
CREATE TABLE IF NOT EXISTS lapinsus_reports LIKE reports;

INSERT IGNORE INTO lapinhar_reports SELECT * FROM reports WHERE report_type = 'lapinhar';
INSERT IGNORE INTO lapinsus_reports SELECT * FROM reports WHERE report_type = 'lapinsus';

ALTER TABLE report_attachments DROP FOREIGN KEY fk_attachments_report;
ALTER TABLE document_number_reservations DROP FOREIGN KEY fk_number_reservation_report;

DROP TABLE reports;

ALTER TABLE lapinsus_reports AUTO_INCREMENT = 1000000;

CREATE VIEW reports AS
SELECT * FROM lapinhar_reports
UNION ALL
SELECT * FROM lapinsus_reports;
