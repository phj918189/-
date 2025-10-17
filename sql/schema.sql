CREATE TABLE IF NOT EXISTS samples(
  id INTEGER PRIMARY KEY,
  sample_no TEXT, site_name TEXT, collected_at TEXT,
  kind TEXT, item TEXT, status TEXT,
  uniq_key TEXT UNIQUE, raw_path TEXT, created_at TEXT
);
CREATE TABLE IF NOT EXISTS researchers(
  id INTEGER PRIMARY KEY, name TEXT, email TEXT,
  slack_id TEXT, skills_json TEXT, active INTEGER
);
CREATE TABLE IF NOT EXISTS assignments(
  id INTEGER PRIMARY KEY, sample_no TEXT, item TEXT,
  researcher TEXT, assigned_at TEXT, method TEXT
);
