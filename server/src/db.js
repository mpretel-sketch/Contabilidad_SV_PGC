import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { DatabaseSync } from "node:sqlite";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const configuredDataDir = process.env.DATA_DIR ? path.resolve(process.env.DATA_DIR) : null;
const configuredDbPath = process.env.DB_PATH ? path.resolve(process.env.DB_PATH) : null;
const dataDir = configuredDataDir || path.resolve(__dirname, "../data");
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });

const dbPath = configuredDbPath || path.join(dataDir, "contabilidad.db");
const dbParentDir = path.dirname(dbPath);
if (!fs.existsSync(dbParentDir)) fs.mkdirSync(dbParentDir, { recursive: true });
const db = new DatabaseSync(dbPath);

db.exec(`
CREATE TABLE IF NOT EXISTS periods (
  period_key TEXT PRIMARY KEY,
  year INTEGER NOT NULL,
  month INTEGER NOT NULL,
  filename TEXT,
  exchange_rate REAL NOT NULL,
  uploaded_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS period_rows (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  period_key TEXT NOT NULL,
  row_id TEXT NOT NULL,
  sort_order INTEGER NOT NULL,
  code TEXT NOT NULL,
  name TEXT,
  sid REAL NOT NULL,
  sia REAL NOT NULL,
  cargos REAL NOT NULL,
  abonos REAL NOT NULL,
  sfd REAL NOT NULL,
  sfa REAL NOT NULL,
  is_new INTEGER NOT NULL DEFAULT 0,
  exclude_from_analysis INTEGER NOT NULL DEFAULT 0,
  UNIQUE(period_key, row_id)
);

CREATE TABLE IF NOT EXISTS period_manual_mappings (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  period_key TEXT NOT NULL,
  row_id TEXT NOT NULL,
  pgc TEXT,
  pgc_name TEXT,
  grupo TEXT,
  subgrupo TEXT,
  UNIQUE(period_key, row_id)
);
`);

const pad2 = (value) => String(value).padStart(2, "0");
export const buildPeriodKey = (year, month) => `${year}-${pad2(month)}`;

export function savePeriodData({ year, month, filename, exchangeRate, rows, manualMappings = {} }) {
  const periodKey = buildPeriodKey(year, month);
  const uploadedAt = new Date().toISOString();

  const upsertPeriod = db.prepare(`
    INSERT INTO periods(period_key, year, month, filename, exchange_rate, uploaded_at)
    VALUES(?, ?, ?, ?, ?, ?)
    ON CONFLICT(period_key) DO UPDATE SET
      filename = excluded.filename,
      exchange_rate = excluded.exchange_rate,
      uploaded_at = excluded.uploaded_at
  `);

  const deleteRows = db.prepare(`DELETE FROM period_rows WHERE period_key = ?`);
  const deleteMappings = db.prepare(`DELETE FROM period_manual_mappings WHERE period_key = ?`);

  const insertRow = db.prepare(`
    INSERT INTO period_rows(
      period_key, row_id, sort_order, code, name, sid, sia, cargos, abonos, sfd, sfa, is_new, exclude_from_analysis
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const insertMapping = db.prepare(`
    INSERT INTO period_manual_mappings(period_key, row_id, pgc, pgc_name, grupo, subgrupo)
    VALUES (?, ?, ?, ?, ?, ?)
  `);

  try {
    db.exec("BEGIN");
    upsertPeriod.run(periodKey, year, month, filename || null, exchangeRate, uploadedAt);
    deleteRows.run(periodKey);
    deleteMappings.run(periodKey);

    rows.forEach((row, index) => {
      insertRow.run(
        periodKey,
        row._rowId,
        index,
        String(row.code ?? ""),
        String(row.name ?? ""),
        Number(row.sid || 0),
        Number(row.sia || 0),
        Number(row.cargos || 0),
        Number(row.abonos || 0),
        Number(row.sfd || 0),
        Number(row.sfa || 0),
        row._isNew ? 1 : 0,
        row._excludeFromAnalysis ? 1 : 0
      );
    });

    Object.entries(manualMappings).forEach(([rowId, mapping]) => {
      if (!mapping) return;
      insertMapping.run(
        periodKey,
        rowId,
        mapping.pgc || null,
        mapping.pgcName || null,
        mapping.grupo || null,
        mapping.subgrupo || null
      );
    });
    db.exec("COMMIT");
  } catch (error) {
    db.exec("ROLLBACK");
    throw error;
  }
  return { periodKey, uploadedAt };
}

export function listPeriods() {
  const stmt = db.prepare(`
    SELECT
      p.period_key,
      p.year,
      p.month,
      p.filename,
      p.exchange_rate,
      p.uploaded_at,
      (SELECT COUNT(1) FROM period_rows r WHERE r.period_key = p.period_key) AS row_count
    FROM periods p
    ORDER BY p.year DESC, p.month DESC
  `);
  return stmt.all();
}

export function loadPeriodData({ year, month }) {
  const periodKey = buildPeriodKey(year, month);
  const periodStmt = db.prepare(`SELECT * FROM periods WHERE period_key = ?`);
  const period = periodStmt.get(periodKey);
  if (!period) return null;

  const rowsStmt = db.prepare(`
    SELECT row_id, code, name, sid, sia, cargos, abonos, sfd, sfa, is_new, exclude_from_analysis
    FROM period_rows
    WHERE period_key = ?
    ORDER BY sort_order ASC
  `);

  const mapStmt = db.prepare(`
    SELECT row_id, pgc, pgc_name, grupo, subgrupo
    FROM period_manual_mappings
    WHERE period_key = ?
  `);

  const rows = rowsStmt.all(periodKey).map((row) => ({
    _rowId: row.row_id,
    _isNew: Boolean(row.is_new),
    _excludeFromAnalysis: Boolean(row.exclude_from_analysis),
    code: row.code,
    name: row.name,
    sid: Number(row.sid || 0),
    sia: Number(row.sia || 0),
    cargos: Number(row.cargos || 0),
    abonos: Number(row.abonos || 0),
    sfd: Number(row.sfd || 0),
    sfa: Number(row.sfa || 0)
  }));

  const manualMappings = {};
  mapStmt.all(periodKey).forEach((row) => {
    manualMappings[row.row_id] = {
      pgc: row.pgc || "",
      pgcName: row.pgc_name || "",
      grupo: row.grupo || "Sin clasificar",
      subgrupo: row.subgrupo || "Sin clasificar"
    };
  });

  return {
    period: {
      year: period.year,
      month: period.month,
      filename: period.filename,
      exchangeRate: Number(period.exchange_rate || 0),
      uploadedAt: period.uploaded_at,
      periodKey
    },
    rows,
    manualMappings
  };
}

export function deletePeriodData({ year, month }) {
  const periodKey = buildPeriodKey(year, month);
  const periodExistsStmt = db.prepare(`SELECT period_key FROM periods WHERE period_key = ?`);
  const found = periodExistsStmt.get(periodKey);
  if (!found) return false;

  const deleteRows = db.prepare(`DELETE FROM period_rows WHERE period_key = ?`);
  const deleteMappings = db.prepare(`DELETE FROM period_manual_mappings WHERE period_key = ?`);
  const deletePeriod = db.prepare(`DELETE FROM periods WHERE period_key = ?`);

  try {
    db.exec("BEGIN");
    deleteRows.run(periodKey);
    deleteMappings.run(periodKey);
    deletePeriod.run(periodKey);
    db.exec("COMMIT");
    return true;
  } catch (error) {
    db.exec("ROLLBACK");
    throw error;
  }
}

export function loadYearPeriodsUntilMonth({ year, month, excludePeriodKey = null }) {
  const stmt = db.prepare(`
    SELECT year, month
    FROM periods
    WHERE year = ? AND month <= ?
    ORDER BY month ASC
  `);

  const rows = stmt.all(Number(year), Number(month));
  const out = [];
  for (const row of rows) {
    const periodKey = buildPeriodKey(row.year, row.month);
    if (excludePeriodKey && periodKey === excludePeriodKey) continue;
    const payload = loadPeriodData({ year: row.year, month: row.month });
    if (payload) out.push(payload);
  }
  return out;
}
