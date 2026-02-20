import express from "express";
import cors from "cors";
import multer from "multer";
import * as XLSX from "xlsx";
import {
  ACCOUNT_MAPPING,
  SAMPLE_ROWS,
  buildExportWorkbook,
  convertRows,
  normalizeRows,
  parseWorkbookBuffer
} from "./conversionEngine.js";
import {
  deletePeriodData,
  listPeriods,
  loadPeriodData,
  loadYearPeriodsUntilMonth,
  savePeriodData
} from "./db.js";

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
const PORT = process.env.PORT || 4000;
const pad2 = (value) => String(value).padStart(2, "0");

app.use(cors());
app.use(express.json({ limit: "20mb" }));

app.get("/api/health", (_req, res) => {
  res.json({ ok: true, service: "sv-pgc-converter", at: new Date().toISOString() });
});

app.get("/api/sample", (_req, res) => {
  res.json({ rows: SAMPLE_ROWS });
});

app.get("/api/mapping/meta", (_req, res) => {
  const groups = {};
  for (const value of Object.values(ACCOUNT_MAPPING)) {
    groups[value.grupo] = (groups[value.grupo] || 0) + 1;
  }
  res.json({ totalMappings: Object.keys(ACCOUNT_MAPPING).length, groups });
});

app.get("/api/periods", (_req, res) => {
  const periods = listPeriods();
  res.json({ periods });
});

app.get("/api/periods/check/:year/:month", (req, res) => {
  const year = Number(req.params.year);
  const month = Number(req.params.month);
  if (!year || !month) {
    return res.status(400).json({ error: "Debes indicar mes y anio validos." });
  }
  const payload = loadPeriodData({ year, month });
  if (!payload) {
    return res.json({
      exists: false,
      period: { year, month },
      rowCount: 0,
      mappingCount: 0,
      uploadedAt: null
    });
  }
  return res.json({
    exists: true,
    period: { year, month },
    rowCount: payload.rows.length,
    mappingCount: Object.keys(payload.manualMappings || {}).length,
    uploadedAt: payload.period?.uploadedAt || null,
    filename: payload.period?.filename || null
  });
});

function getPreviousPeriod(month, year) {
  if (month === 1) return { month: 12, year: year - 1 };
  return { month: month - 1, year };
}

function differenceRowsByPreviousMonth(currentRows, previousRows) {
  const prevByCode = new Map(previousRows.map((row) => [String(row.code || "").trim(), row]));
  const currentCodes = new Set(currentRows.map((row) => String(row.code || "").trim()));
  const numericFields = ["sid", "sia", "cargos", "abonos", "sfd", "sfa"];

  const subtract = (a, b) => Number(a || 0) - Number(b || 0);
  const hasAnyAmount = (row) => numericFields.some((field) => Math.abs(Number(row[field] || 0)) > 1e-9);

  const diffRows = currentRows.map((current, index) => {
    const code = String(current.code || "").trim();
    const previous = prevByCode.get(code) || {};
    const next = {
      _rowId: current._rowId || `row-${index + 1}`,
      _isNew: Boolean(current._isNew),
      _excludeFromAnalysis: Boolean(current._excludeFromAnalysis),
      code,
      name: current.name || previous.name || ""
    };
    for (const field of numericFields) {
      next[field] = subtract(current[field], previous[field]);
    }
    return next;
  });

  for (const previous of previousRows) {
    const code = String(previous.code || "").trim();
    if (!code || currentCodes.has(code)) continue;
    const prevOnly = {
      _rowId: `prev-${code}`,
      _isNew: false,
      _excludeFromAnalysis: false,
      code,
      name: previous.name || "",
      sid: -Number(previous.sid || 0),
      sia: -Number(previous.sia || 0),
      cargos: -Number(previous.cargos || 0),
      abonos: -Number(previous.abonos || 0),
      sfd: -Number(previous.sfd || 0),
      sfa: -Number(previous.sfa || 0)
    };
    if (hasAnyAmount(prevOnly)) diffRows.push(prevOnly);
  }

  return diffRows;
}

function aggregateRowsByCode(rows) {
  const map = new Map();
  const numericFields = ["sid", "sia", "cargos", "abonos", "sfd", "sfa"];

  for (const row of rows) {
    const code = String(row.code || "").trim();
    if (!code) continue;
    if (!map.has(code)) {
      map.set(code, {
        _rowId: `ytd-${code}`,
        _isNew: false,
        _excludeFromAnalysis: false,
        code,
        name: row.name || "",
        sid: 0,
        sia: 0,
        cargos: 0,
        abonos: 0,
        sfd: 0,
        sfa: 0
      });
    }
    const target = map.get(code);
    if (row.name) target.name = row.name;
    for (const field of numericFields) {
      target[field] += Number(row[field] || 0);
    }
  }
  return Array.from(map.values());
}

function collectManualMappingsByCode(rows, manualMappings = {}) {
  const out = new Map();
  for (const row of rows) {
    const code = String(row?.code || "").trim();
    if (!code) continue;
    const manual = manualMappings[row._rowId];
    if (!manual || typeof manual !== "object") continue;
    const hasData = Boolean(manual.pgc || manual.pgcName || manual.grupo || manual.subgrupo);
    if (!hasData) continue;
    out.set(code, {
      pgc: manual.pgc || "",
      pgcName: manual.pgcName || "",
      grupo: manual.grupo || "Sin clasificar",
      subgrupo: manual.subgrupo || "Sin clasificar"
    });
  }
  return out;
}

function buildYtdConversion({
  year,
  month,
  exchangeRate,
  currentRows = [],
  currentManualMappings = {},
  excludeCurrentPeriodInDb = false
}) {
  const currentPeriodKey = `${year}-${pad2(month)}`;
  const savedPayloads = loadYearPeriodsUntilMonth({
    year,
    month,
    excludePeriodKey: excludeCurrentPeriodInDb ? currentPeriodKey : null
  });

  const ytdSourceRows = [];
  const ytdMappingsByCode = new Map();
  for (const payload of savedPayloads) {
    ytdSourceRows.push(...payload.rows);
    const payloadMap = collectManualMappingsByCode(payload.rows, payload.manualMappings);
    for (const [code, mapping] of payloadMap.entries()) ytdMappingsByCode.set(code, mapping);
  }
  ytdSourceRows.push(...currentRows);
  const currentMap = collectManualMappingsByCode(currentRows, currentManualMappings);
  for (const [code, mapping] of currentMap.entries()) ytdMappingsByCode.set(code, mapping);

  const ytdRows = aggregateRowsByCode(ytdSourceRows);
  const ytdManualMappings = {};
  for (const row of ytdRows) {
    const code = String(row.code || "").trim();
    if (!code) continue;
    const mapping = ytdMappingsByCode.get(code);
    if (mapping) ytdManualMappings[row._rowId] = mapping;
  }

  const ytdConversion = convertRows(ytdRows, exchangeRate, ytdManualMappings, { month, year, scope: "YTD" });
  return { ytdRows, ytdConversion };
}

const MONTH_NAMES = [
  "Enero",
  "Febrero",
  "Marzo",
  "Abril",
  "Mayo",
  "Junio",
  "Julio",
  "Agosto",
  "Septiembre",
  "Octubre",
  "Noviembre",
  "Diciembre"
];

function comparePgcCode(a, b) {
  const an = Number(String(a || "").replace(/\D/g, ""));
  const bn = Number(String(b || "").replace(/\D/g, ""));
  if (Number.isFinite(an) && Number.isFinite(bn) && an !== bn) return an - bn;
  return String(a || "").localeCompare(String(b || ""));
}

function extractPnlLineMap(conversion) {
  const map = new Map();
  for (const section of Object.values(conversion?.pnl?.sections || {})) {
    for (const item of section.items || []) {
      const key = `${item.pgcCode}__${item.pgcName}`;
      if (!map.has(key)) {
        map.set(key, {
          cuenta: item.pgcCode,
          descripcion: item.pgcName,
          mxn: Number(item.totalMXN || 0),
          eur: Number(item.totalEUR || 0)
        });
      } else {
        const current = map.get(key);
        current.mxn += Number(item.totalMXN || 0);
        current.eur += Number(item.totalEUR || 0);
      }
    }
  }
  return map;
}

function buildPnlWorkbookByMonths({
  year,
  month,
  exchangeRate,
  currentRows,
  currentManualMappings
}) {
  const wb = XLSX.utils.book_new();
  const monthConversions = new Map();

  for (let m = 1; m <= 12; m += 1) {
    if (m > month) break;
    if (m === month) {
      monthConversions.set(
        m,
        convertRows(currentRows, exchangeRate, currentManualMappings, { year, month: m })
      );
      continue;
    }
    const payload = loadPeriodData({ year, month: m });
    if (!payload) continue;
    monthConversions.set(
      m,
      convertRows(payload.rows, payload.period.exchangeRate || exchangeRate, payload.manualMappings, { year, month: m })
    );
  }

  const lines = new Map();
  for (const [m, conversion] of monthConversions.entries()) {
    const lineMap = extractPnlLineMap(conversion);
    for (const line of lineMap.values()) {
      const key = `${line.cuenta}__${line.descripcion}`;
      if (!lines.has(key)) {
        lines.set(key, { cuenta: line.cuenta, descripcion: line.descripcion, monthsMXN: {}, monthsEUR: {} });
      }
      const target = lines.get(key);
      target.monthsMXN[m] = Number(line.mxn || 0);
      target.monthsEUR[m] = Number(line.eur || 0);
    }
  }

  const baseRows = Array.from(lines.values()).sort((a, b) => {
    const byCode = comparePgcCode(a.cuenta, b.cuenta);
    if (byCode !== 0) return byCode;
    return a.descripcion.localeCompare(b.descripcion);
  });

  const buildRowsForCurrency = (field) => {
    const rows = baseRows.map((line) => {
      const row = {
        Cuenta: line.cuenta,
        Descripcion: line.descripcion
      };
      for (let m = 1; m <= 12; m += 1) {
        row[MONTH_NAMES[m - 1]] = Number(line[field][m] || 0);
      }
      return row;
    });

    const totalRow = { Cuenta: "TOTAL", Descripcion: "Resultado del periodo (PyG)" };
    for (let m = 1; m <= 12; m += 1) {
      const conversion = monthConversions.get(m);
      totalRow[MONTH_NAMES[m - 1]] = Number(
        field === "monthsEUR"
          ? conversion?.pnl?.resultadoAntesImpuestosEur || 0
          : conversion?.pnl?.resultadoAntesImpuestosMx || 0
      );
    }
    rows.push(totalRow);
    return rows;
  };

  const wsMXN = XLSX.utils.json_to_sheet(buildRowsForCurrency("monthsMXN"));
  XLSX.utils.book_append_sheet(wb, wsMXN, "P&L_BASE");

  const wsEUR = XLSX.utils.json_to_sheet(buildRowsForCurrency("monthsEUR"));
  XLSX.utils.book_append_sheet(wb, wsEUR, "P&L_EUR");

  return wb;
}

app.get("/api/periods/:year/:month", (req, res) => {
  const year = Number(req.params.year);
  const month = Number(req.params.month);
  const payload = loadPeriodData({ year, month });
  if (!payload) {
    return res.status(404).json({ error: "No existe informacion para ese periodo." });
  }

  const conversion = convertRows(
    payload.rows,
    payload.period.exchangeRate || 1.0,
    payload.manualMappings,
    { month, year }
  );
  const { ytdConversion } = buildYtdConversion({
    year,
    month,
    exchangeRate: payload.period.exchangeRate || 1.0,
    currentRows: [],
    currentManualMappings: {},
    excludeCurrentPeriodInDb: false
  });

  return res.json({
    ...conversion,
    ytdConversion,
    sourceRows: payload.rows,
    manualMappings: payload.manualMappings,
    storage: payload.period
  });
});

app.delete("/api/periods/:year/:month", (req, res) => {
  try {
    const year = Number(req.params.year);
    const month = Number(req.params.month);
    if (!year || !month) {
      return res.status(400).json({ error: "Debes indicar mes y anio validos." });
    }
    const deleted = deletePeriodData({ year, month });
    if (!deleted) {
      return res.status(404).json({ error: "No existe el periodo a eliminar." });
    }
    return res.json({ ok: true, deleted: true, period: { year, month } });
  } catch (error) {
    return res.status(500).json({ error: error.message || "No se pudo eliminar el periodo." });
  }
});

app.post("/api/periods/save", (req, res) => {
  try {
    const {
      rows = [],
      exchangeRate = 1.0,
      manualMappings = {},
      period = {}
    } = req.body || {};

    const year = Number(period.year);
    const month = Number(period.month);
    if (!year || !month) {
      return res.status(400).json({ error: "Debes indicar mes y anio del periodo." });
    }

    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({ error: "Debes enviar filas contables en 'rows'." });
    }

    const normalized = normalizeRows(rows);
    const conversion = convertRows(normalized.rows, Number(exchangeRate) || 1.0, manualMappings, { month, year });
    const hasUnmapped = conversion.metadata.unmappedCount > 0;
    const hasAnalyzedRows = conversion.metadata.analyzedRowCount > 0;
    const trialDiff = Math.abs(conversion.validations.trialBalanceFinalDifference);
    const canPersist = hasAnalyzedRows && !hasUnmapped && trialDiff <= 0.01;

    if (!canPersist) {
      return res.status(400).json({
        error: "No se puede guardar: hay partidas sin mapear o descuadre en balanza final.",
        validations: {
          analyzedRowCount: conversion.metadata.analyzedRowCount,
          unmappedCount: conversion.metadata.unmappedCount,
          trialBalanceFinalDifference: conversion.validations.trialBalanceFinalDifference
        }
      });
    }

    savePeriodData({
      year,
      month,
      rows: normalized.rows,
      manualMappings,
      exchangeRate: Number(exchangeRate) || 1.0,
      filename: period.filename || "manual-save"
    });
    const { ytdConversion } = buildYtdConversion({
      year,
      month,
      exchangeRate: Number(exchangeRate) || 1.0,
      currentRows: normalized.rows,
      currentManualMappings: manualMappings,
      excludeCurrentPeriodInDb: true
    });
    return res.json({
      ...conversion,
      ytdConversion,
      sourceRows: normalized.rows,
      manualMappings,
      storage: { year, month, exchangeRate: Number(exchangeRate) || 1.0 }
    });
  } catch (error) {
    return res.status(500).json({ error: error.message || "No se pudo guardar el periodo." });
  }
});

app.post("/api/periods/upload", upload.single("file"), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No se encontro ningun archivo. Usa el campo 'file'." });
    }

    const month = Number(req.body.month);
    const year = Number(req.body.year);
    if (!year || !month) {
      return res.status(400).json({ error: "Debes seleccionar mes y anio." });
    }

    const exchangeRate = Number(req.body.exchangeRate) || 1.0;
    const rows = parseWorkbookBuffer(req.file.buffer);
    const previousPeriod = getPreviousPeriod(month, year);
    const previousPayload = loadPeriodData(previousPeriod);
    const rowsForPeriod = previousPayload
      ? differenceRowsByPreviousMonth(rows, previousPayload.rows)
      : rows;

    const conversion = convertRows(rowsForPeriod, exchangeRate, {}, { month, year });
    const { ytdConversion } = buildYtdConversion({
      year,
      month,
      exchangeRate,
      currentRows: rowsForPeriod,
      currentManualMappings: {},
      excludeCurrentPeriodInDb: true
    });
    return res.json({
      ...conversion,
      ytdConversion,
      sourceRows: rowsForPeriod,
      manualMappings: {},
      storage: {
        year,
        month,
        exchangeRate,
        filename: req.file.originalname,
        persisted: false,
        differenceApplied: Boolean(previousPayload),
        basePeriod: previousPayload ? previousPeriod : null
      }
    });
  } catch (error) {
    return res.status(400).json({ error: error.message || "No se pudo leer o guardar el archivo." });
  }
});

app.post("/api/convert", (req, res) => {
  try {
    const { rows = [], exchangeRate = 1.0, manualMappings = {}, period = null } = req.body || {};
    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({ error: "Debes enviar filas contables en 'rows'." });
    }

    const normalized = normalizeRows(rows);
    const year = Number(period?.year);
    const month = Number(period?.month);
    const conversion = convertRows(normalized.rows, Number(exchangeRate) || 1.0, manualMappings, period);
    const ytdConversion =
      year && month
        ? buildYtdConversion({
            year,
            month,
            exchangeRate: Number(exchangeRate) || 1.0,
            currentRows: normalized.rows,
            currentManualMappings: manualMappings,
            excludeCurrentPeriodInDb: true
          }).ytdConversion
        : null;
    return res.json({ ...conversion, ytdConversion });
  } catch (error) {
    return res.status(500).json({ error: error.message || "No se pudo convertir la informacion." });
  }
});

app.post("/api/export", (req, res) => {
  try {
    const { rows = [], exchangeRate = 1.0, manualMappings = {}, period = null } = req.body || {};
    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({ error: "Debes enviar filas contables en 'rows'." });
    }

    const normalized = normalizeRows(rows);
    const conversion = convertRows(normalized.rows, Number(exchangeRate) || 1.0, manualMappings, period);
    const workbook = buildExportWorkbook(conversion);
    const outputBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    const periodLabel =
      period?.month && period?.year
        ? `${String(period.year)}-${String(period.month).padStart(2, "0")}`
        : new Date().toISOString().slice(0, 10);

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename=conversion_pgc_${periodLabel}.xlsx`);
    return res.send(outputBuffer);
  } catch (error) {
    return res.status(500).json({ error: error.message || "No se pudo generar el archivo Excel." });
  }
});

app.post("/api/export-pnl", (req, res) => {
  try {
    const { rows = [], exchangeRate = 1.0, manualMappings = {}, period = null } = req.body || {};
    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({ error: "Debes enviar filas contables en 'rows'." });
    }

    const normalized = normalizeRows(rows);
    const exportYear = Number(period?.year) || new Date().getFullYear();
    const exportMonth = Number(period?.month) || new Date().getMonth() + 1;
    const workbook = buildPnlWorkbookByMonths({
      year: exportYear,
      month: exportMonth,
      exchangeRate: Number(exchangeRate) || 1.0,
      currentRows: normalized.rows,
      currentManualMappings: manualMappings
    });
    const outputBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    const periodLabel =
      period?.month && period?.year
        ? `${String(period.year)}-${String(period.month).padStart(2, "0")}`
        : new Date().toISOString().slice(0, 10);

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename=pnl_esp_${periodLabel}.xlsx`);
    return res.send(outputBuffer);
  } catch (error) {
    return res.status(500).json({ error: error.message || "No se pudo generar el archivo PyG." });
  }
});

app.listen(PORT, () => {
  // eslint-disable-next-line no-console
  console.log(`Server running on http://localhost:${PORT}`);
});
