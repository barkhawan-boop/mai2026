CREATE TABLE IF NOT EXISTS report_status (
  inspector_id INTEGER NOT NULL,
  activity_year INTEGER NOT NULL,
  activity_month INTEGER NOT NULL CHECK (activity_month BETWEEN 1 AND 12),
  sent_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (inspector_id, activity_year, activity_month)
);
