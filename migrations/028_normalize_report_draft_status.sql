UPDATE lapinhar_reports
SET status = 'selesai'
WHERE status = 'draft'
  AND COALESCE(title, '') <> ''
  AND COALESCE(facts, '') <> '';

UPDATE lapinsus_reports
SET status = 'selesai'
WHERE status = 'draft'
  AND COALESCE(title, '') <> ''
  AND COALESCE(facts, '') <> '';
