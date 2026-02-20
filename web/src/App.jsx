import { Fragment, useEffect, useMemo, useRef, useState } from "react";

const TABS = [
  { key: "partidas", label: "Partidas (detalle)" },
  { key: "mapping", label: "Mapeo" },
  { key: "balance", label: "Balance" },
  { key: "pnl", label: "P&G" },
  { key: "control", label: "Control" }
];

const GROUP_OPTIONS = [
  "Sin clasificar",
  "Activo No Corriente",
  "Activo Corriente",
  "Patrimonio Neto",
  "Pasivo No Corriente",
  "Pasivo Corriente",
  "Ingresos",
  "Ingresos Financieros",
  "Gastos",
  "Gastos Financieros"
];

const SUBGROUP_OPTIONS = [
  "Sin clasificar",
  "Efectivo y equivalentes",
  "Deudores comerciales",
  "Otros deudores",
  "Periodificaciones",
  "Administraciones Publicas",
  "Inmovilizado material",
  "Amortizacion acumulada",
  "Inversiones financieras LP",
  "Acreedores comerciales",
  "Otros acreedores",
  "Deudas empresas grupo CP",
  "Otros pasivos",
  "Deudas a LP",
  "Fondos propios",
  "Importe neto cifra negocios",
  "Otros ingresos de explotacion",
  "Gastos de personal",
  "Servicios exteriores",
  "Tributos",
  "Amortizaciones",
  "Gastos excepcionales",
  "Resultado financiero",
  "Otros resultados"
];

const MONTHS = [
  { value: 1, label: "Enero" },
  { value: 2, label: "Febrero" },
  { value: 3, label: "Marzo" },
  { value: 4, label: "Abril" },
  { value: 5, label: "Mayo" },
  { value: 6, label: "Junio" },
  { value: 7, label: "Julio" },
  { value: 8, label: "Agosto" },
  { value: 9, label: "Septiembre" },
  { value: 10, label: "Octubre" },
  { value: 11, label: "Noviembre" },
  { value: 12, label: "Diciembre" }
];

const NUMERIC_FIELDS = ["sid", "sia", "cargos", "abonos", "sfd", "sfa"];

const fmt = (n) =>
  new Intl.NumberFormat("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n || 0);

const fmtEur = (n) =>
  new Intl.NumberFormat("es-ES", { style: "currency", currency: "EUR", minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n || 0);

const toDisplay = (value, showEur) => (showEur ? fmtEur(value) : fmt(value));

const rowFromAny = (row, index = 0) => ({
  _rowId: String(row._rowId || row.rowId || row.id || `row-${index + 1}`),
  _isNew: Boolean(row._isNew),
  _excludeFromAnalysis: Boolean(row._excludeFromAnalysis),
  code: String(row.code || ""),
  name: String(row.name || ""),
  sid: Number(row.sid || 0),
  sia: Number(row.sia || 0),
  cargos: Number(row.cargos || 0),
  abonos: Number(row.abonos || 0),
  sfd: Number(row.sfd || 0),
  sfa: Number(row.sfa || 0)
});

const parseNum = (value) => {
  const n = Number.parseFloat(String(value));
  return Number.isFinite(n) ? n : 0;
};

async function api(path, options = {}) {
  const fallbackProdApi = "https://back-contabilidad-sv-pgc-onrender-com.onrender.com";
  const configuredBase = import.meta.env.VITE_API_BASE_URL || (import.meta.env.PROD ? fallbackProdApi : "");
  const normalizedBase = String(configuredBase || "").trim();
  const withProtocol =
    normalizedBase && !/^https?:\/\//i.test(normalizedBase)
      ? `https://${normalizedBase}`
      : normalizedBase;
  const baseUrl = String(withProtocol || "")
    .replace(/\/api\/?$/, "")
    .replace(/\/$/, "");
  const safePath = path.startsWith("/") ? path : `/${path}`;
  const response = await fetch(`${baseUrl}${safePath}`, options);
  if (!response.ok) {
    const raw = await response.text().catch(() => "");
    let parsed = null;
    try {
      parsed = raw ? JSON.parse(raw) : null;
    } catch {
      parsed = null;
    }
    throw new Error(parsed?.error || `Error API (${response.status})`);
  }
  return response;
}

async function safeJson(response) {
  const raw = await response.text().catch(() => "");
  if (!raw) return {};
  try {
    return JSON.parse(raw);
  } catch {
    throw new Error("La API devolvio una respuesta no JSON");
  }
}

export default function App() {
  const now = new Date();
  const [tab, setTab] = useState("partidas");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [exchangeRate, setExchangeRate] = useState(1.0);
  const [showEur, setShowEur] = useState(false);
  const [conversion, setConversion] = useState(null);
  const [sourceRows, setSourceRows] = useState([]);
  const [manualMappings, setManualMappings] = useState({});
  const [search, setSearch] = useState("");
  const [groupFilter, setGroupFilter] = useState("Todos");
  const [expanded, setExpanded] = useState("");
  const [mappingMeta, setMappingMeta] = useState(null);
  const [detailSearch, setDetailSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("todos");
  const [collapsedPnlSections, setCollapsedPnlSections] = useState({});
  const [periodMonth, setPeriodMonth] = useState(() => Number(localStorage.getItem("pt_period_month")) || now.getMonth() + 1);
  const [periodYear, setPeriodYear] = useState(() => Number(localStorage.getItem("pt_period_year")) || now.getFullYear());
  const [periods, setPeriods] = useState([]);
  const [showAllPeriods, setShowAllPeriods] = useState(false);
  const [pendingDeletePeriodKey, setPendingDeletePeriodKey] = useState("");
  const [periodCheck, setPeriodCheck] = useState({
    exists: false,
    rowCount: 0,
    mappingCount: 0,
    uploadedAt: null,
    filename: null
  });
  const [viewScope, setViewScope] = useState("month");
  const detailTableRef = useRef(null);
  const detailTopScrollRef = useRef(null);
  const syncScrollLockRef = useRef(false);
  const [detailScrollTop, setDetailScrollTop] = useState(0);
  const [detailViewportHeight, setDetailViewportHeight] = useState(560);
  const [detailTableWidth, setDetailTableWidth] = useState(2860);
  const [detailScrollLeft, setDetailScrollLeft] = useState(0);
  const [detailMaxScrollLeft, setDetailMaxScrollLeft] = useState(0);

  useEffect(() => {
    localStorage.setItem("pt_period_month", String(periodMonth));
    localStorage.setItem("pt_period_year", String(periodYear));
  }, [periodMonth, periodYear]);

  const refreshPeriods = async () => {
    const response = await api("/api/periods");
    const payload = await safeJson(response);
    setPeriods(payload.periods || []);
  };

  const checkPeriodStorage = async (month = periodMonth, year = periodYear) => {
    try {
      const response = await api(`/api/periods/check/${year}/${month}`);
      const payload = await safeJson(response);
      setPeriodCheck({
        exists: Boolean(payload.exists),
        rowCount: Number(payload.rowCount || 0),
        mappingCount: Number(payload.mappingCount || 0),
        uploadedAt: payload.uploadedAt || null,
        filename: payload.filename || null
      });
    } catch {
      setPeriodCheck({
        exists: false,
        rowCount: 0,
        mappingCount: 0,
        uploadedAt: null,
        filename: null
      });
    }
  };

  const deletePeriod = async (month, year) => {
    setLoading(true);
    setError("");
    try {
      await api(`/api/periods/${year}/${month}`, { method: "DELETE" });
      await refreshPeriods();
      await checkPeriodStorage(periodMonth, periodYear);
      if (Number(periodMonth) === Number(month) && Number(periodYear) === Number(year)) {
        setConversion(null);
        setSourceRows([]);
        setManualMappings({});
      }
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
      setPendingDeletePeriodKey("");
    }
  };

  const hydrateState = (payload) => {
    const rows = (payload.sourceRows || payload.convertedData || []).map(rowFromAny);
    setSourceRows(rows);
    setManualMappings(payload.manualMappings || {});
    setExchangeRate(Number(payload.metadata?.exchangeRate || payload.storage?.exchangeRate || 1.0));
    setConversion(payload);
    setExpanded("");
    setSearch("");
    setDetailSearch("");
    setGroupFilter("Todos");
  };

  const loadPeriod = async (month = periodMonth, year = periodYear) => {
    setLoading(true);
    setError("");
    try {
      const response = await api(`/api/periods/${year}/${month}`);
      const payload = await safeJson(response);
      hydrateState(payload);
      await checkPeriodStorage(month, year);
    } catch (e) {
      const msg = String(e.message || "");
      if (msg.includes("No existe informacion") || msg.includes("Error API (404)")) {
        setConversion(null);
        setSourceRows([]);
        setManualMappings({});
      } else {
        setError(msg);
      }
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    const bootstrap = async () => {
      setLoading(true);
      setError("");
      try {
        const [metaRes] = await Promise.all([api("/api/mapping/meta").then((r) => safeJson(r))]);
        setMappingMeta(metaRes);
        await refreshPeriods();
        await checkPeriodStorage(periodMonth, periodYear);
      } catch (e) {
        setError(e.message);
      } finally {
        setLoading(false);
      }
    };

    bootstrap();
  }, []);

  useEffect(() => {
    checkPeriodStorage(periodMonth, periodYear);
  }, [periodMonth, periodYear]);

  const savePeriod = async () => {
    if (!sourceRows.length) {
      setError("No hay lineas para guardar.");
      return;
    }
    setLoading(true);
    setError("");
    try {
      const response = await api("/api/periods/save", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          rows: sourceRows,
          manualMappings,
          exchangeRate,
          period: { month: periodMonth, year: periodYear }
        })
      });
      const payload = await safeJson(response);
      hydrateState(payload);
      await refreshPeriods();
      await checkPeriodStorage(periodMonth, periodYear);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  const onUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const formData = new FormData();
    formData.append("file", file);
    formData.append("exchangeRate", String(exchangeRate));
    formData.append("month", String(periodMonth));
    formData.append("year", String(periodYear));

    setLoading(true);
    setError("");
    try {
      const response = await api("/api/periods/upload", { method: "POST", body: formData });
      const payload = await safeJson(response);
      hydrateState(payload);
      await refreshPeriods();
      await checkPeriodStorage(periodMonth, periodYear);
      setTab("partidas");
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  const onExport = async () => {
    if (!sourceRows.length) return;
    setLoading(true);
    setError("");
    try {
      const response = await api("/api/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          rows: sourceRows,
          manualMappings,
          exchangeRate,
          period: { month: periodMonth, year: periodYear }
        })
      });
      const blob = await response.blob();
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = `conversion_nif_a_pgc_${periodYear}-${String(periodMonth).padStart(2, "0")}.xlsx`;
      link.click();
      URL.revokeObjectURL(link.href);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  const onExportPnl = async () => {
    if (!sourceRows.length) return;
    setLoading(true);
    setError("");
    try {
      const response = await api("/api/export-pnl", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          rows: sourceRows,
          manualMappings,
          exchangeRate,
          period: { month: periodMonth, year: periodYear }
        })
      });
      const blob = await response.blob();
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = `pnl_esp_${periodYear}-${String(periodMonth).padStart(2, "0")}.xlsx`;
      link.click();
      URL.revokeObjectURL(link.href);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  const recalculateConversion = async () => {
    if (!sourceRows.length) return;
    setLoading(true);
    setError("");
    try {
      const response = await api("/api/convert", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          rows: sourceRows,
          manualMappings,
          exchangeRate,
          period: { month: periodMonth, year: periodYear }
        })
      });
      const payload = await safeJson(response);
      setConversion(payload);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  const updateRow = (rowId, field, value) => {
    setSourceRows((rows) =>
      rows.map((row) => {
        if (row._rowId !== rowId) return row;
        const next = { ...row };
        next[field] = NUMERIC_FIELDS.includes(field) ? parseNum(value) : value;
        return next;
      })
    );
  };

  const removeRow = (rowId) => {
    setSourceRows((rows) => rows.filter((row) => row._rowId !== rowId));
    setManualMappings((prev) => {
      const next = { ...prev };
      delete next[rowId];
      return next;
    });
  };

  const addRow = () => {
    const rowId = `manual-${Date.now()}`;
    setSourceRows((rows) => [
      {
        _rowId: rowId,
        _isNew: true,
        _excludeFromAnalysis: false,
        code: "",
        name: "",
        sid: 0,
        sia: 0,
        cargos: 0,
        abonos: 0,
        sfd: 0,
        sfa: 0
      },
      ...rows
    ]);
  };

  const updateManualMapping = (rowId, field, value) => {
    setManualMappings((prev) => {
      const base = prev[rowId] || { pgc: "", pgcName: "", grupo: "Sin clasificar", subgrupo: "Sin clasificar" };
      return { ...prev, [rowId]: { ...base, [field]: value } };
    });
  };

  const clearManualMapping = (rowId) => {
    setManualMappings((prev) => {
      const next = { ...prev };
      delete next[rowId];
      return next;
    });
  };

  const reportData = useMemo(() => {
    if (!conversion) return null;
    if (viewScope === "ytd" && conversion.ytdConversion) return conversion.ytdConversion;
    return conversion;
  }, [conversion, viewScope]);

  const groups = useMemo(() => {
    if (!reportData?.pgcAggregated) return ["Todos"];
    return ["Todos", ...new Set(reportData.pgcAggregated.map((item) => item.grupo))];
  }, [reportData]);

  const filtered = useMemo(() => {
    if (!reportData?.pgcAggregated) return [];
    const query = search.trim().toLowerCase();

    return reportData.pgcAggregated.filter((item) => {
      if (groupFilter !== "Todos" && item.grupo !== groupFilter) return false;
      if (!query) return true;
      const inHeader = item.pgcCode.toLowerCase().includes(query) || item.pgcName.toLowerCase().includes(query);
      const inDetails = item.details.some((d) => d.code.toLowerCase().includes(query) || d.name.toLowerCase().includes(query));
      return inHeader || inDetails;
    });
  }, [reportData, groupFilter, search]);

  const filteredRows = useMemo(() => {
    const q = detailSearch.trim().toLowerCase();
    return sourceRows.filter((row) => row.code.toLowerCase().includes(q) || row.name.toLowerCase().includes(q));
  }, [sourceRows, detailSearch]);

  const convertedByRowId = useMemo(() => {
    const map = new Map();
    (conversion?.convertedData || []).forEach((row) => map.set(row._rowId, row));
    return map;
  }, [conversion]);

  const filteredRowsByStatus = useMemo(() => {
    return filteredRows.filter((row) => {
      const c = convertedByRowId.get(row._rowId);
      const isSummary = Boolean(c?.isSummaryLine);
      const isMapped = Boolean(c && c.pgcCode !== "SIN MAPEO");
      if (statusFilter === "sumatorias") return isSummary;
      if (statusFilter === "mapeadas") return !isSummary && isMapped;
      if (statusFilter === "sin-mapear") return !isSummary && !isMapped;
      return true;
    });
  }, [filteredRows, convertedByRowId, statusFilter]);

  const virtualRows = useMemo(() => {
    const ROW_HEIGHT = 50;
    const OVERSCAN = 10;
    const total = filteredRowsByStatus.length;
    const viewport = Math.max(320, detailViewportHeight || 560);
    const visibleCount = Math.max(18, Math.ceil(viewport / ROW_HEIGHT) + OVERSCAN * 2);
    const start = Math.max(0, Math.floor(detailScrollTop / ROW_HEIGHT) - OVERSCAN);
    const end = Math.min(total, start + visibleCount);
    return {
      rows: filteredRowsByStatus.slice(start, end),
      start,
      end,
      total,
      topSpacer: start * ROW_HEIGHT,
      bottomSpacer: Math.max(0, (total - end) * ROW_HEIGHT)
    };
  }, [filteredRowsByStatus, detailScrollTop, detailViewportHeight]);

  const coverageClass = (coverage = 0) => {
    if (coverage >= 95) return "ok";
    if (coverage >= 80) return "warn";
    return "bad";
  };

  const pnlSectionEntries = useMemo(
    () => Object.entries(reportData?.pnl?.sections || {}),
    [reportData]
  );

  const setAllPnlCollapsed = (collapsed) => {
    const next = {};
    pnlSectionEntries.forEach(([section]) => {
      next[section] = collapsed;
    });
    setCollapsedPnlSections(next);
  };

  const togglePnlSection = (section) => {
    setCollapsedPnlSections((prev) => ({ ...prev, [section]: !prev[section] }));
  };

  const checks = useMemo(() => {
    const unmapped = conversion?.metadata?.unmappedCount || 0;
    const analyzed = conversion?.metadata?.analyzedRowCount || 0;
    const trialDiff = Math.abs(conversion?.validations?.trialBalanceFinalDifference || 0);
    const canSave = conversion && sourceRows.length > 0 && analyzed > 0 && unmapped === 0 && trialDiff <= 0.01;
    return { unmapped, analyzed, trialDiff, canSave };
  }, [conversion, sourceRows]);

  const visiblePeriods = useMemo(() => {
    if (showAllPeriods) return periods;
    return periods.slice(0, 8);
  }, [periods, showAllPeriods]);

  const balanceDrift = useMemo(() => {
    if (!conversion) return 0;
    return Math.abs(conversion.balanceSheet?.adjustedDifferenceMXN || conversion.balanceSheet?.differenceMXN || 0);
  }, [conversion]);

  const balanceRightWithAdjustment = useMemo(() => {
    if (!reportData) return 0;
    return showEur
      ? reportData.balanceSheet.adjustedTotalPasivoPNEUR
      : reportData.balanceSheet.adjustedTotalPasivoPNMXN;
  }, [reportData, showEur]);

  const diffInfo = conversion?.storage?.differenceApplied
    ? `Diferencia aplicada contra ${String(conversion.storage.basePeriod?.month || "").padStart(2, "0")}/${conversion.storage.basePeriod?.year || ""}`
    : "";

  useEffect(() => {
    if (tab !== "partidas") return;
    setDetailScrollTop(0);
    setDetailScrollLeft(0);
    if (detailTableRef.current) detailTableRef.current.scrollTop = 0;
    if (detailTableRef.current) detailTableRef.current.scrollLeft = 0;
    if (detailTopScrollRef.current) detailTopScrollRef.current.scrollLeft = 0;
  }, [detailSearch, statusFilter, tab]);

  useEffect(() => {
    const updateWidths = () => {
      const wrapEl = detailTableRef.current;
      const tableEl = wrapEl?.querySelector(".detail-table");
      if (!tableEl || !wrapEl) return;
      const nextWidth = Math.ceil(tableEl.scrollWidth || 2860);
      setDetailTableWidth(nextWidth);
      const nextMaxLeft = Math.max(0, wrapEl.scrollWidth - wrapEl.clientWidth);
      setDetailMaxScrollLeft(nextMaxLeft);
      setDetailScrollLeft(Math.min(wrapEl.scrollLeft || 0, nextMaxLeft));
    };
    updateWidths();
    window.addEventListener("resize", updateWidths);
    return () => window.removeEventListener("resize", updateWidths);
  }, [tab, sourceRows.length, statusFilter, detailSearch]);

  const onTopHorizontalScroll = (event) => {
    if (!detailTableRef.current) return;
    if (syncScrollLockRef.current) return;
    syncScrollLockRef.current = true;
    detailTableRef.current.scrollLeft = event.currentTarget.scrollLeft || 0;
    requestAnimationFrame(() => {
      syncScrollLockRef.current = false;
    });
  };

  const onDetailTableScroll = (event) => {
    setDetailScrollTop(event.currentTarget.scrollTop || 0);
    setDetailViewportHeight(event.currentTarget.clientHeight || 560);
    const maxLeft = Math.max(0, event.currentTarget.scrollWidth - event.currentTarget.clientWidth);
    setDetailMaxScrollLeft(maxLeft);
    setDetailScrollLeft(Math.min(event.currentTarget.scrollLeft || 0, maxLeft));
    if (detailTopScrollRef.current) {
      if (syncScrollLockRef.current) return;
      syncScrollLockRef.current = true;
      detailTopScrollRef.current.scrollLeft = event.currentTarget.scrollLeft || 0;
      requestAnimationFrame(() => {
        syncScrollLockRef.current = false;
      });
    }
  };

  const scrollDetailHorizontally = (delta) => {
    const wrapEl = detailTableRef.current;
    if (!wrapEl) return;
    const nextLeft = Math.max(0, Math.min(wrapEl.scrollLeft + delta, wrapEl.scrollWidth - wrapEl.clientWidth));
    wrapEl.scrollLeft = nextLeft;
    setDetailScrollLeft(nextLeft);
  };

  const onDetailRangeScroll = (event) => {
    const wrapEl = detailTableRef.current;
    if (!wrapEl) return;
    const nextLeft = Number(event.target.value || 0);
    wrapEl.scrollLeft = nextLeft;
    if (detailTopScrollRef.current) detailTopScrollRef.current.scrollLeft = nextLeft;
    setDetailScrollLeft(nextLeft);
  };

  return (
    <div className="app-shell">
      <header className="topbar topbar-ops">
        <div className="ops-group">
          <span className="ops-label">Periodo</span>
          <div className="period-control">
            <select value={periodMonth} onChange={(e) => setPeriodMonth(Number(e.target.value))}>
              {MONTHS.map((m) => (
                <option key={m.value} value={m.value}>{m.label}</option>
              ))}
            </select>
            <input type="number" min="2000" max="2100" value={periodYear} onChange={(e) => setPeriodYear(Number(e.target.value) || now.getFullYear())} />
            <button type="button" className="btn subtle" onClick={() => loadPeriod(periodMonth, periodYear)}>Cargar</button>
          </div>
        </div>

        <div className="ops-group">
          <span className="ops-label">Carga</span>
          <label className="file-input">
            Subir archivo
            <input type="file" accept=".xlsx,.xls,.csv" onChange={onUpload} />
          </label>
        </div>

        <div className="ops-group">
          <span className="ops-label">Moneda / Tipo cambio</span>
          <div className="ops-line">
            <button type="button" className="btn subtle" onClick={() => setShowEur((s) => !s)}>{showEur ? "EUR" : "BASE"}</button>
            <div className="rate-control">
              <span>TC EUR/EUR</span>
              <input type="number" step="0.001" value={exchangeRate} onChange={(e) => setExchangeRate(Number.parseFloat(e.target.value) || 1.0)} />
            </div>
          </div>
        </div>

        <div className="ops-group ops-group-wide">
          <span className="ops-label">Acciones</span>
          <div className="ops-line">
            <button type="button" className="btn subtle" onClick={recalculateConversion} disabled={!sourceRows.length}>
              Recalcular
            </button>
            <button type="button" className="btn success" onClick={savePeriod} disabled={!checks.canSave}>
              Guardar periodo
            </button>
            <div className="btn-group">
              <button type="button" className="btn subtle btn-export" onClick={onExportPnl} disabled={!sourceRows.length}>P&L XLSX</button>
              <button type="button" className="btn subtle btn-export" onClick={onExport} disabled={!sourceRows.length}>Exportar XLSX</button>
            </div>
          </div>
        </div>

        <div className="ops-group">
          <span className="ops-label">Vista</span>
          <div className="view-switch">
            <button
              type="button"
              className={viewScope === "month" ? "mini-btn active-view" : "mini-btn"}
              onClick={() => setViewScope("month")}
            >
              Mes subido
            </button>
            <button
              type="button"
              className={viewScope === "ytd" ? "mini-btn active-view" : "mini-btn"}
              onClick={() => setViewScope("ytd")}
              disabled={!conversion?.ytdConversion}
            >
              YTD
            </button>
          </div>
        </div>
      </header>

      {error && <div className="alert bad">{error}</div>}
      {loading && <div className="alert">Procesando...</div>}
      {diffInfo && <div className="alert">{diffInfo}</div>}

      <main className="workspace">
        <aside className="sidebar">
          <div className="sidebar-sticky">
            <section className="side-brand">
              <div className="brand-mark">P</div>
              <div>
                <h4>SNC PT - PGC ESP</h4>
                <p>Panel mensual</p>
              </div>
            </section>
            <nav className="side-tabs">
              {TABS.map((item) => (
                <button type="button" key={item.key} className={tab === item.key ? "side-tab active" : "side-tab"} onClick={() => setTab(item.key)}>
                  {item.label}
                </button>
              ))}
            </nav>
            {conversion && (
              <section className="side-card">
                <h4>Estado guardado</h4>
                <p>{checks.canSave ? "Listo para guardar" : "Pendiente de corregir"}</p>
                <p>Sin mapear: <strong>{checks.unmapped}</strong></p>
                <p>Dif. balanza final: <strong>{fmt(checks.trialDiff)}</strong></p>
              </section>
            )}
            <section className="side-card">
              <h4>Periodos guardados</h4>
              <div className="period-list">
                {periods.length === 0 && <span className="muted-inline">Sin periodos guardados.</span>}
                {periods.length > 0 && (
                  <p className="muted-inline">Mostrando {visiblePeriods.length} de {periods.length}</p>
                )}
                {visiblePeriods.map((p) => (
                  <div key={p.period_key} className="period-item">
                    <button type="button" className="mini-btn period-load-btn" onClick={() => { setPeriodMonth(Number(p.month)); setPeriodYear(Number(p.year)); loadPeriod(Number(p.month), Number(p.year)); }}>
                      {String(p.month).padStart(2, "0")}/{p.year} · {p.row_count}
                    </button>
                    <button
                      type="button"
                      className="mini-btn danger"
                      onClick={() => {
                        if (pendingDeletePeriodKey !== p.period_key) {
                          setPendingDeletePeriodKey(p.period_key);
                          return;
                        }
                        deletePeriod(Number(p.month), Number(p.year));
                      }}
                    >
                      {pendingDeletePeriodKey === p.period_key ? "Confirmar borrado" : "Eliminar"}
                    </button>
                    {pendingDeletePeriodKey === p.period_key && (
                      <button type="button" className="mini-btn" onClick={() => setPendingDeletePeriodKey("")}>
                        Cancelar
                      </button>
                    )}
                  </div>
                ))}
                {periods.length > 8 && (
                  <button type="button" className="mini-btn period-toggle-btn" onClick={() => setShowAllPeriods((s) => !s)}>
                    {showAllPeriods ? "Ver menos" : "Ver todos"}
                  </button>
                )}
              </div>
            </section>
            {mappingMeta && (
              <section className="side-card">
                <h4>Mapeos base</h4>
                <p>Total: <strong>{mappingMeta.totalMappings}</strong></p>
              </section>
            )}
          </div>
        </aside>
        <section className="workspace-main">
          <section className="cards">
          <article className="card"><span className="label">{viewScope === "ytd" ? "Lineas YTD" : "Lineas mes"}</span><strong>{reportData?.metadata?.rowCount || sourceRows.length}</strong></article>
          <article className="card"><span className="label">Lineas analizadas</span><strong>{reportData?.metadata?.analyzedRowCount || 0}</strong></article>
          <article className="card"><span className="label">Sumatorias excluidas</span><strong>{reportData?.metadata?.summaryExcludedCount || 0}</strong></article>
          <article className="card"><span className="label">Sin mapear</span><strong className={(reportData?.metadata?.unmappedCount || 0) === 0 ? "ok" : "bad"}>{reportData?.metadata?.unmappedCount || 0}</strong></article>
          <article className="card"><span className="label">Cobertura mapeo</span><strong className={coverageClass(reportData?.metadata?.mappedCoveragePct || 0)}>{fmt(reportData?.metadata?.mappedCoveragePct || 0)}%</strong></article>
          <article className="card"><span className="label">Descuadre final</span><strong className={Math.abs(reportData?.validations?.trialBalanceFinalDifference || 0) <= 0.01 ? "ok" : "bad"}>{fmt(Math.abs(reportData?.validations?.trialBalanceFinalDifference || 0))}</strong></article>
          </section>
        {conversion && tab === "partidas" && (
          <section>
            <div className="partidas-toolbar">
              <button type="button" className="btn subtle" onClick={addRow}>Anadir partida</button>
              <button type="button" className="btn subtle" onClick={recalculateConversion}>Recalcular conversion</button>
            </div>
            <div className="controls">
              <input type="text" placeholder="Buscar partida por cuenta o descripcion" value={detailSearch} onChange={(e) => setDetailSearch(e.target.value)} />
              <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
                <option value="todos">Todos los estados</option>
                <option value="sin-mapear">Sin mapear</option>
                <option value="mapeadas">Mapeadas</option>
                <option value="sumatorias">Sumatorias excluidas</option>
              </select>
            </div>
            <div className="table-info-bar">
              <span>Filas visibles: {virtualRows.rows.length}</span>
              <span>Total filtradas: {virtualRows.total}</span>
              <span>Render virtual activo</span>
            </div>
            <div className="table-nav">
              <button type="button" className="mini-btn" onClick={() => scrollDetailHorizontally(-320)}>←</button>
              <input
                className="table-nav-range"
                type="range"
                min={0}
                max={Math.max(0, detailMaxScrollLeft)}
                step={1}
                value={Math.min(detailScrollLeft, Math.max(0, detailMaxScrollLeft))}
                onChange={onDetailRangeScroll}
              />
              <button type="button" className="mini-btn" onClick={() => scrollDetailHorizontally(320)}>→</button>
            </div>
            <div className="table-scroll-top" ref={detailTopScrollRef} onScroll={onTopHorizontalScroll}>
              <div className="table-scroll-inner" style={{ width: `${detailTableWidth}px` }} />
            </div>
            <div className="table-wrap" ref={detailTableRef} onScroll={onDetailTableScroll}>
              <table className="detail-table">
                <colgroup>
                  <col style={{ width: "150px" }} />
                  <col style={{ width: "260px" }} />
                  <col style={{ width: "115px" }} />
                  <col style={{ width: "115px" }} />
                  <col style={{ width: "125px" }} />
                  <col style={{ width: "125px" }} />
                  <col style={{ width: "120px" }} />
                  <col style={{ width: "120px" }} />
                  <col style={{ width: "115px" }} />
                  <col style={{ width: "115px" }} />
                  <col style={{ width: "120px" }} />
                  <col style={{ width: "125px" }} />
                  <col style={{ width: "110px" }} />
                  <col style={{ width: "260px" }} />
                  <col style={{ width: "95px" }} />
                  <col style={{ width: "230px" }} />
                  <col style={{ width: "155px" }} />
                  <col style={{ width: "170px" }} />
                  <col style={{ width: "150px" }} />
                </colgroup>
                <thead>
                  <tr>
                    <th className="sticky-col sticky-col-1">Cuenta PT</th>
                    <th className="sticky-col sticky-col-2">Descripcion</th>
                    <th className="right">SID</th><th className="right">SIA</th><th className="right">Cargos</th><th className="right">Abonos</th>
                    <th className="right">Suma Debe</th><th className="right">Suma Haber</th><th className="right">SFD</th><th className="right">SFA</th><th className="right">Saldo Neto</th>
                    <th>Estado</th><th>PGC asignado</th><th>Nombre asignado</th><th>PGC</th><th>Nombre PGC</th><th>Grupo</th><th>Subgrupo</th><th className="right">Acciones</th>
                  </tr>
                </thead>
                <tbody>
                  {virtualRows.topSpacer > 0 && (
                    <tr aria-hidden="true">
                      <td colSpan={19} style={{ height: `${virtualRows.topSpacer}px`, padding: 0, borderBottom: "none" }} />
                    </tr>
                  )}
                  {virtualRows.rows.map((row) => {
                    const manual = manualMappings[row._rowId] || { pgc: "", pgcName: "", grupo: "Sin clasificar", subgrupo: "Sin clasificar" };
                    const converted = convertedByRowId.get(row._rowId);
                    const isSummary = Boolean(converted?.isSummaryLine);
                    const isMapped = Boolean(converted && converted.pgcCode !== "SIN MAPEO");
                    const sumaDebe = Number(row.sid || 0) + Number(row.cargos || 0);
                    const sumaHaber = Number(row.sia || 0) + Number(row.abonos || 0);
                    const saldoNeto = Number(row.sfd || 0) - Number(row.sfa || 0);
                    return (
                      <tr key={row._rowId} className={isSummary ? "summary-row" : isMapped ? "mapped-row" : "unmapped-row"}>
                        <td className="sticky-col sticky-col-1"><input value={row.code} onChange={(e) => updateRow(row._rowId, "code", e.target.value)} className="cell-input" /></td>
                        <td className="sticky-col sticky-col-2"><input value={row.name} onChange={(e) => updateRow(row._rowId, "name", e.target.value)} className="cell-input" /></td>
                        <td className="right"><input value={row.sid} onChange={(e) => updateRow(row._rowId, "sid", e.target.value)} className="cell-input num" /></td>
                        <td className="right"><input value={row.sia} onChange={(e) => updateRow(row._rowId, "sia", e.target.value)} className="cell-input num" /></td>
                        <td className="right"><input value={row.cargos} onChange={(e) => updateRow(row._rowId, "cargos", e.target.value)} className="cell-input num" /></td>
                        <td className="right"><input value={row.abonos} onChange={(e) => updateRow(row._rowId, "abonos", e.target.value)} className="cell-input num" /></td>
                        <td className="right">{fmt(sumaDebe)}</td>
                        <td className="right">{fmt(sumaHaber)}</td>
                        <td className="right"><input value={row.sfd} onChange={(e) => updateRow(row._rowId, "sfd", e.target.value)} className="cell-input num" /></td>
                        <td className="right"><input value={row.sfa} onChange={(e) => updateRow(row._rowId, "sfa", e.target.value)} className="cell-input num" /></td>
                        <td className="right">{fmt(saldoNeto)}</td>
                        <td><span className={isSummary ? "badge badge-warn" : isMapped ? "badge badge-ok" : "badge badge-bad"}>{isSummary ? "Sumatoria excluida" : isMapped ? "Mapeada" : "Sin mapear"}</span></td>
                        <td>{converted?.pgcCode || "SIN MAPEO"}</td>
                        <td>{converted?.pgcName || "Sin equivalencia PGC"}</td>
                        <td><input value={manual.pgc} onChange={(e) => updateManualMapping(row._rowId, "pgc", e.target.value)} className="cell-input" /></td>
                        <td><input value={manual.pgcName} onChange={(e) => updateManualMapping(row._rowId, "pgcName", e.target.value)} className="cell-input" /></td>
                        <td><select value={manual.grupo} onChange={(e) => updateManualMapping(row._rowId, "grupo", e.target.value)} className="cell-select">{GROUP_OPTIONS.map((opt) => <option key={opt} value={opt}>{opt}</option>)}</select></td>
                        <td><select value={manual.subgrupo} onChange={(e) => updateManualMapping(row._rowId, "subgrupo", e.target.value)} className="cell-select">{SUBGROUP_OPTIONS.map((opt) => <option key={opt} value={opt}>{opt}</option>)}</select></td>
                        <td className="right nowrap-actions">
                          <button className="mini-btn" type="button" onClick={() => clearManualMapping(row._rowId)}>Reset map</button>
                          <button className="mini-btn danger" type="button" onClick={() => removeRow(row._rowId)}>Eliminar</button>
                        </td>
                      </tr>
                    );
                  })}
                  {virtualRows.bottomSpacer > 0 && (
                    <tr aria-hidden="true">
                      <td colSpan={19} style={{ height: `${virtualRows.bottomSpacer}px`, padding: 0, borderBottom: "none" }} />
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </section>
        )}

        {reportData && tab === "mapping" && (
          <section>
            <div className="controls">
              <select value={groupFilter} onChange={(e) => setGroupFilter(e.target.value)}>{groups.map((g) => <option key={g} value={g}>{g}</option>)}</select>
              <input type="text" placeholder="Buscar por cuenta o nombre" value={search} onChange={(e) => setSearch(e.target.value)} />
            </div>
            <div className="table-wrap">
              <table>
                <thead><tr><th>PGC</th><th>Nombre</th><th>Grupo</th><th className="right">Total</th><th className="right">Detalle</th></tr></thead>
                <tbody>
                  {filtered.map((row) => (
                    <Fragment key={row.pgcCode}>
                      <tr onClick={() => setExpanded(expanded === row.pgcCode ? "" : row.pgcCode)}>
                        <td>{row.pgcCode}</td><td>{row.pgcName}</td><td>{row.grupo}</td><td className="right">{toDisplay(showEur ? row.totalEUR : row.totalMXN, showEur)}</td><td className="right">{row.details.length}</td>
                      </tr>
                      {expanded === row.pgcCode && row.details.map((detail) => (
                        <tr className="child" key={`${row.pgcCode}-${detail._rowId}`}>
                          <td>{detail.code}</td><td>{detail.name}</td><td>{detail.subgrupo}</td><td className="right">{toDisplay(showEur ? detail.displayEUR : detail.displayMXN, showEur)}</td>
                          <td className="right">SID {fmt(detail.sid)} | SIA {fmt(detail.sia)} | C {fmt(detail.cargos)} | A {fmt(detail.abonos)} | SFD {fmt(detail.sfd)} | SFA {fmt(detail.sfa)}</td>
                        </tr>
                      ))}
                    </Fragment>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        )}

        {reportData && tab === "balance" && (
          <section className="grid-2">
            <article className="panel">
              <h3>Activo</h3>
              {["Activo No Corriente", "Activo Corriente"].map((group) => (
                <div key={group} className="block">
                  <h4>{group}</h4>
                  {reportData.balanceSheet.groups[group].items.map((item) => <p key={`${group}-${item.pgcCode}`}><span>{item.pgcCode} {item.pgcName}</span><span>{toDisplay(showEur ? item.totalEUR : item.totalMXN, showEur)}</span></p>)}
                  <p className="total"><span>Total {group}</span><span>{toDisplay(showEur ? reportData.balanceSheet.groups[group].totalEUR : reportData.balanceSheet.groups[group].totalMXN, showEur)}</span></p>
                </div>
              ))}
              <p className="grand-total"><span>Total Activo</span><span>{toDisplay(showEur ? reportData.balanceSheet.totalActivoEUR : reportData.balanceSheet.totalActivoMXN, showEur)}</span></p>
            </article>
            <article className="panel">
              <h3>Patrimonio neto y pasivo</h3>
              {["Patrimonio Neto", "Pasivo No Corriente", "Pasivo Corriente"].map((group) => (
                <div key={group} className="block">
                  <h4>{group}</h4>
                  {reportData.balanceSheet.groups[group].items.map((item) => <p key={`${group}-${item.pgcCode}`}><span>{item.pgcCode} {item.pgcName}</span><span>{toDisplay(showEur ? item.totalEUR : item.totalMXN, showEur)}</span></p>)}
                  <p className="total"><span>Total {group}</span><span>{toDisplay(showEur ? reportData.balanceSheet.groups[group].totalEUR : reportData.balanceSheet.groups[group].totalMXN, showEur)}</span></p>
                </div>
              ))}
              <p className="grand-total"><span>Total PN + Pasivo</span><span>{toDisplay(showEur ? reportData.balanceSheet.totalPasivoPNEUR : reportData.balanceSheet.totalPasivoPNMXN, showEur)}</span></p>
              {reportData.balanceSheet.autoResultLine && (
                <p className="grand-total">
                  <span>129 Resultado del periodo (ajuste tecnico)</span>
                  <span>{toDisplay(showEur ? reportData.balanceSheet.autoResultLine.totalEUR : reportData.balanceSheet.autoResultLine.totalMXN, showEur)}</span>
                </p>
              )}
              <p className="grand-total"><span>Total PN + Pasivo (ajustado)</span><span>{toDisplay(balanceRightWithAdjustment, showEur)}</span></p>
            </article>
          </section>
        )}

        {reportData && tab === "pnl" && (
          <section className="panel">
            <h3>Cuenta de Perdidas y Ganancias</h3>
            <div className="pnl-toolbar">
              <button type="button" className="mini-btn" onClick={() => setAllPnlCollapsed(false)}>Expandir todo</button>
              <button type="button" className="mini-btn" onClick={() => setAllPnlCollapsed(true)}>Colapsar todo</button>
            </div>
            {pnlSectionEntries.map(([section, content]) => (
              <div className="block" key={section}>
                <button type="button" className="section-toggle" onClick={() => togglePnlSection(section)}>
                  <span>{collapsedPnlSections[section] ? "▸" : "▾"} {section}</span>
                  <span>Subtotal: {toDisplay(showEur ? content.totalEUR : content.totalMXN, showEur)}</span>
                </button>
                {!collapsedPnlSections[section] && (
                  <>
                    {content.items.map((item) => <p key={`${section}-${item.pgcCode}`}><span>{item.pgcCode} {item.pgcName}</span><span>{toDisplay(showEur ? item.totalEUR : item.totalMXN, showEur)}</span></p>)}
                    <p className="total"><span>Subtotal</span><span>{toDisplay(showEur ? content.totalEUR : content.totalMXN, showEur)}</span></p>
                  </>
                )}
              </div>
            ))}
            <div className="result-card">
              <span className="label">Resultado del periodo</span>
              <strong className={(showEur ? reportData.pnl.resultadoAntesImpuestosEur : reportData.pnl.resultadoAntesImpuestosMx) < 0 ? "bad" : "ok"}>
                {toDisplay(showEur ? reportData.pnl.resultadoAntesImpuestosEur : reportData.pnl.resultadoAntesImpuestosMx, showEur)}
              </strong>
            </div>
          </section>
        )}

        {reportData && tab === "control" && (
          <section className="panel">
            <h3>Controles de calidad</h3>
            <p>Guardado en BBDD: <strong className={periodCheck.exists ? "ok" : "bad"}>{periodCheck.exists ? "OK" : "NO"}</strong></p>
            <p>Periodo seleccionado: <strong>{String(periodMonth).padStart(2, "0")}/{periodYear}</strong></p>
            <p>Filas guardadas: <strong>{periodCheck.rowCount}</strong></p>
            <p>Mapeos manuales guardados: <strong>{periodCheck.mappingCount}</strong></p>
            <p>Ultima carga guardada: <strong>{periodCheck.uploadedAt ? new Date(periodCheck.uploadedAt).toLocaleString("es-ES") : "-"}</strong></p>
            <p>Lineas analizadas: <strong>{checks.analyzed}</strong></p>
            <p>Sin mapear: <strong>{checks.unmapped}</strong></p>
            <p>Diferencia balanza inicial: <strong>{fmt(reportData.validations.trialBalanceInitialDifference)}</strong></p>
            <p>Diferencia balanza final: <strong>{fmt(reportData.validations.trialBalanceFinalDifference)}</strong></p>
            <p>Diferencia balance PGC: <strong>{fmt(reportData.balanceSheet.differenceMXN)}</strong></p>
            <p>Diferencia balance PGC ajustado: <strong>{fmt(balanceDrift)}</strong></p>
            <p>Cuentas sin mapeo (lineas analizadas): <strong>{reportData.metadata.unmappedCount}</strong></p>
            <p><strong>{checks.canSave ? "Listo para guardar en BBDD" : "No se puede guardar hasta dejar sin mapear=0 y balanza final cuadrada"}</strong></p>
          </section>
        )}

        {mappingMeta && reportData && (
          <footer className="footer">
            <span>Mapeos cargados: {mappingMeta.totalMappings}</span>
            <span>Periodo: {String(periodMonth).padStart(2, "0")}/{periodYear}</span>
            <span>Vista: <strong>{viewScope === "ytd" ? "YTD" : "Mes subido"}</strong></span>
            <span>
              Total PYG: <strong>{toDisplay(showEur ? reportData.pnl.resultadoAntesImpuestosEur : reportData.pnl.resultadoAntesImpuestosMx, showEur)}</strong>
            </span>
            <span>Generado: {new Date(reportData.metadata.generatedAt).toLocaleString("es-ES")}</span>
          </footer>
        )}
        </section>
      </main>
    </div>
  );
}
