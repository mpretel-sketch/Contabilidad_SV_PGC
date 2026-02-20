import * as XLSX from "xlsx";

export const ACCOUNT_MAPPING = {
  "101": { pgc: "570", pgcName: "Caja, euros", grupo: "Activo Corriente", subgrupo: "Efectivo y equivalentes" },
  "102-001": { pgc: "572", pgcName: "Bancos e instituciones de credito c/c vista, euros", grupo: "Activo Corriente", subgrupo: "Efectivo y equivalentes" },
  "102-002": { pgc: "573", pgcName: "Bancos e instituciones de credito c/c vista, moneda extranjera", grupo: "Activo Corriente", subgrupo: "Efectivo y equivalentes" },
  "104-001": { pgc: "430", pgcName: "Clientes", grupo: "Activo Corriente", subgrupo: "Deudores comerciales" },
  "104-002": { pgc: "4304", pgcName: "Clientes, moneda extranjera", grupo: "Activo Corriente", subgrupo: "Deudores comerciales" },
  "108-001": { pgc: "460", pgcName: "Anticipos de remuneraciones", grupo: "Activo Corriente", subgrupo: "Otros deudores" },
  "108-004": { pgc: "440", pgcName: "Deudores", grupo: "Activo Corriente", subgrupo: "Otros deudores" },
  "110": { pgc: "480", pgcName: "Gastos anticipados", grupo: "Activo Corriente", subgrupo: "Periodificaciones" },
  "111": { pgc: "4709", pgcName: "H.P. deudora por devolucion de impuestos", grupo: "Activo Corriente", subgrupo: "Administraciones Publicas" },
  "112": { pgc: "473", pgcName: "H.P. retenciones y pagos a cuenta", grupo: "Activo Corriente", subgrupo: "Administraciones Publicas" },
  "115": { pgc: "4720", pgcName: "H.P. IVA soportado", grupo: "Activo Corriente", subgrupo: "Administraciones Publicas" },
  "116": { pgc: "4720", pgcName: "H.P. IVA soportado (pte. de pago)", grupo: "Activo Corriente", subgrupo: "Administraciones Publicas" },
  "117": { pgc: "407", pgcName: "Anticipos a proveedores", grupo: "Activo Corriente", subgrupo: "Deudores comerciales" },
  "122": { pgc: "218", pgcName: "Elementos de transporte", grupo: "Activo No Corriente", subgrupo: "Inmovilizado material" },
  "123": { pgc: "216", pgcName: "Mobiliario", grupo: "Activo No Corriente", subgrupo: "Inmovilizado material" },
  "124": { pgc: "217", pgcName: "Equipos para procesos de informacion", grupo: "Activo No Corriente", subgrupo: "Inmovilizado material" },
  "135-003": { pgc: "2818", pgcName: "Amort. acum. elementos de transporte", grupo: "Activo No Corriente", subgrupo: "Amortizacion acumulada" },
  "135-004": { pgc: "2816", pgcName: "Amort. acum. mobiliario", grupo: "Activo No Corriente", subgrupo: "Amortizacion acumulada" },
  "135-005": { pgc: "2817", pgcName: "Amort. acum. equipos proceso informacion", grupo: "Activo No Corriente", subgrupo: "Amortizacion acumulada" },
  "142": { pgc: "260", pgcName: "Fianzas constituidas a largo plazo", grupo: "Activo No Corriente", subgrupo: "Inversiones financieras LP" },
  "201-001": { pgc: "400", pgcName: "Proveedores", grupo: "Pasivo Corriente", subgrupo: "Acreedores comerciales" },
  "201-002": { pgc: "4004", pgcName: "Proveedores, moneda extranjera", grupo: "Pasivo Corriente", subgrupo: "Acreedores comerciales" },
  "203-003-001": { pgc: "410", pgcName: "Acreedores por prestaciones de servicios", grupo: "Pasivo Corriente", subgrupo: "Otros acreedores" },
  "203-003-002": { pgc: "5530", pgcName: "Socios, c/c (empresas del grupo)", grupo: "Pasivo Corriente", subgrupo: "Deudas empresas grupo CP" },
  "203-003-003": { pgc: "5530", pgcName: "Socios, c/c (empresas del grupo)", grupo: "Pasivo Corriente", subgrupo: "Deudas empresas grupo CP" },
  "204": { pgc: "438", pgcName: "Anticipos de clientes", grupo: "Pasivo Corriente", subgrupo: "Acreedores comerciales" },
  "206": { pgc: "4770", pgcName: "H.P. IVA repercutido", grupo: "Pasivo Corriente", subgrupo: "Administraciones Publicas" },
  "207": { pgc: "4770", pgcName: "H.P. IVA repercutido (pte. cobro)", grupo: "Pasivo Corriente", subgrupo: "Administraciones Publicas" },
  "208-001": { pgc: "4750", pgcName: "H.P. acreedora por IVA", grupo: "Pasivo Corriente", subgrupo: "Administraciones Publicas" },
  "208-003": { pgc: "4752", pgcName: "H.P. acreedora por impuesto sobre sociedades", grupo: "Pasivo Corriente", subgrupo: "Administraciones Publicas" },
  "208-007": { pgc: "476", pgcName: "Organismos de la Seg. Social acreedores", grupo: "Pasivo Corriente", subgrupo: "Administraciones Publicas" },
  "210-001": { pgc: "4751", pgcName: "H.P. acreedora por retenciones practicadas", grupo: "Pasivo Corriente", subgrupo: "Administraciones Publicas" },
  "210-003": { pgc: "4751", pgcName: "H.P. acreedora por retenciones practicadas", grupo: "Pasivo Corriente", subgrupo: "Administraciones Publicas" },
  "210-004": { pgc: "4751", pgcName: "H.P. acreedora por retenciones practicadas", grupo: "Pasivo Corriente", subgrupo: "Administraciones Publicas" },
  "210-008": { pgc: "476", pgcName: "Organismos de la Seg. Social acreedores", grupo: "Pasivo Corriente", subgrupo: "Administraciones Publicas" },
  "210-010": { pgc: "476", pgcName: "Organismos de la Seg. Social acreedores", grupo: "Pasivo Corriente", subgrupo: "Administraciones Publicas" },
  "212": { pgc: "465", pgcName: "Remuneraciones pendientes de pago", grupo: "Pasivo Corriente", subgrupo: "Otros pasivos" },
  "213": { pgc: "171", pgcName: "Deudas a largo plazo", grupo: "Pasivo No Corriente", subgrupo: "Deudas a LP" },
  "301-001": { pgc: "100", pgcName: "Capital social", grupo: "Patrimonio Neto", subgrupo: "Fondos propios" },
  "301-004": { pgc: "1030", pgcName: "Socios por desembolsos no exigidos, capital social", grupo: "Patrimonio Neto", subgrupo: "Fondos propios" },
  "302": { pgc: "112", pgcName: "Reserva legal", grupo: "Patrimonio Neto", subgrupo: "Fondos propios" },
  "303": { pgc: "129", pgcName: "Resultado del ejercicio", grupo: "Patrimonio Neto", subgrupo: "Fondos propios" },
  "304-001": { pgc: "120", pgcName: "Remanente", grupo: "Patrimonio Neto", subgrupo: "Fondos propios" },
  "304-002": { pgc: "121", pgcName: "Resultados negativos de ejercicios anteriores", grupo: "Patrimonio Neto", subgrupo: "Fondos propios" },
  "401": { pgc: "705", pgcName: "Prestaciones de servicios", grupo: "Ingresos", subgrupo: "Importe neto cifra negocios" },
  "402": { pgc: "709", pgcName: "Rappels sobre ventas", grupo: "Ingresos", subgrupo: "Importe neto cifra negocios" },
  "601-000-0001": { pgc: "640", pgcName: "Sueldos y salarios", grupo: "Gastos", subgrupo: "Gastos de personal" },
  "601-000-0002": { pgc: "640", pgcName: "Sueldos y salarios", grupo: "Gastos", subgrupo: "Gastos de personal" },
  "601-000-0003": { pgc: "640", pgcName: "Sueldos y salarios", grupo: "Gastos", subgrupo: "Gastos de personal" },
  "601-000-0004": { pgc: "640", pgcName: "Sueldos y salarios", grupo: "Gastos", subgrupo: "Gastos de personal" },
  "601-000-0005": { pgc: "629", pgcName: "Otros servicios (dietas/viajes)", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0006": { pgc: "640", pgcName: "Sueldos y salarios", grupo: "Gastos", subgrupo: "Gastos de personal" },
  "601-000-0007": { pgc: "622", pgcName: "Comunicaciones", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0008": { pgc: "628", pgcName: "Suministros (agua)", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0009": { pgc: "628", pgcName: "Suministros (electricidad)", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0010": { pgc: "625", pgcName: "Primas de seguros / Vigilancia", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0011": { pgc: "629", pgcName: "Otros servicios (material oficina)", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0013": { pgc: "622", pgcName: "Reparaciones y conservacion", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0014": { pgc: "625", pgcName: "Primas de seguros", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0016": { pgc: "631", pgcName: "Otros tributos", grupo: "Gastos", subgrupo: "Tributos" },
  "601-000-0017": { pgc: "669", pgcName: "Otros gastos financieros (recargos)", grupo: "Gastos", subgrupo: "Gastos financieros" },
  "601-000-0018": { pgc: "629", pgcName: "Otros servicios (cuotas)", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0019": { pgc: "627", pgcName: "Publicidad, propaganda y relaciones publicas", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0020": { pgc: "621", pgcName: "Arrendamientos y canones", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0031": { pgc: "623", pgcName: "Servicios profesionales independientes", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0033": { pgc: "626", pgcName: "Servicios bancarios y similares", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0034": { pgc: "629", pgcName: "Otros servicios (limpieza)", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0035": { pgc: "623", pgcName: "Servicios profesionales independientes", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0036": { pgc: "629", pgcName: "Otros servicios (papeleria)", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0038": { pgc: "624", pgcName: "Transportes", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0039": { pgc: "642", pgcName: "Seg. Social a cargo de la empresa (IMSS)", grupo: "Gastos", subgrupo: "Gastos de personal" },
  "601-000-0040": { pgc: "642", pgcName: "Seg. Social a cargo de la empresa (RCV)", grupo: "Gastos", subgrupo: "Gastos de personal" },
  "601-000-0041": { pgc: "642", pgcName: "Seg. Social a cargo de la empresa (Infonavit)", grupo: "Gastos", subgrupo: "Gastos de personal" },
  "601-000-0044": { pgc: "641", pgcName: "Indemnizaciones (antiguedad)", grupo: "Gastos", subgrupo: "Gastos de personal" },
  "601-000-0045": { pgc: "641", pgcName: "Indemnizaciones", grupo: "Gastos", subgrupo: "Gastos de personal" },
  "601-000-0046": { pgc: "640", pgcName: "Sueldos y salarios (festivos)", grupo: "Gastos", subgrupo: "Gastos de personal" },
  "601-000-0048": { pgc: "629", pgcName: "Otros servicios", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0049": { pgc: "623", pgcName: "Servicios prof. indep. (comisiones reservas)", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0051": { pgc: "649", pgcName: "Otros gastos sociales (comedor)", grupo: "Gastos", subgrupo: "Gastos de personal" },
  "601-000-0052": { pgc: "628", pgcName: "Suministros (combustible)", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0053": { pgc: "623", pgcName: "Servicios prof. independientes (personas morales)", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "601-000-0054": { pgc: "640", pgcName: "Sueldos y salarios (horas extras)", grupo: "Gastos", subgrupo: "Gastos de personal" },
  "601-000-0055": { pgc: "631", pgcName: "Otros tributos (imp. nomina)", grupo: "Gastos", subgrupo: "Tributos" },
  "601-001": { pgc: "678", pgcName: "Gastos excepcionales (no deducibles)", grupo: "Gastos", subgrupo: "Gastos excepcionales" },
  "603-000-0001": { pgc: "623", pgcName: "Servicios prof. indep. (admon)", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "603-000-0002": { pgc: "623", pgcName: "Servicios prof. indep. (PF)", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "603-000-0003": { pgc: "623", pgcName: "Servicios prof. indep. (PM)", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "603-000-0004": { pgc: "623", pgcName: "Servicios profesionales (legal)", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "701": { pgc: "668", pgcName: "Diferencias negativas de cambio", grupo: "Gastos Financieros", subgrupo: "Resultado financiero" },
  "702": { pgc: "768", pgcName: "Diferencias positivas de cambio", grupo: "Ingresos Financieros", subgrupo: "Resultado financiero" },
  "801": { pgc: "678", pgcName: "Gastos excepcionales", grupo: "Gastos", subgrupo: "Otros resultados" },
  "803": { pgc: "678", pgcName: "Gastos excepcionales (no deducibles)", grupo: "Gastos", subgrupo: "Otros resultados" },
  "804-003": { pgc: "6818", pgcName: "Amort. inmovilizado material (transporte)", grupo: "Gastos", subgrupo: "Amortizaciones" },
  "804-004": { pgc: "6816", pgcName: "Amort. inmovilizado material (mobiliario)", grupo: "Gastos", subgrupo: "Amortizaciones" },
  "804-005": { pgc: "6817", pgcName: "Amort. inmovilizado material (eq. informatico)", grupo: "Gastos", subgrupo: "Amortizaciones" },
  "806": { pgc: "759", pgcName: "Ingresos por servicios diversos", grupo: "Ingresos", subgrupo: "Otros ingresos de explotacion" }
};

const PT_INITIAL_PREFIX_MAPPING = {
  "11": { pgc: "570", pgcName: "Caja, euros", grupo: "Activo Corriente", subgrupo: "Efectivo y equivalentes" },
  "12": { pgc: "572", pgcName: "Bancos e instituciones de credito c/c vista, euros", grupo: "Activo Corriente", subgrupo: "Efectivo y equivalentes" },
  "21": { pgc: "430", pgcName: "Clientes", grupo: "Activo Corriente", subgrupo: "Deudores comerciales" },
  "22": { pgc: "400", pgcName: "Proveedores", grupo: "Pasivo Corriente", subgrupo: "Acreedores comerciales" },
  "23": { pgc: "465", pgcName: "Remuneraciones pendientes de pago", grupo: "Pasivo Corriente", subgrupo: "Otros pasivos" },
  "24": { pgc: "4750", pgcName: "H.P. acreedora por IVA", grupo: "Pasivo Corriente", subgrupo: "Administraciones Publicas" },
  "24.1.1": { pgc: "473", pgcName: "H.P. retenciones y pagos a cuenta", grupo: "Activo Corriente", subgrupo: "Administraciones Publicas" },
  "24.2": { pgc: "4720", pgcName: "H.P. IVA soportado", grupo: "Activo Corriente", subgrupo: "Administraciones Publicas" },
  "24.3": { pgc: "4770", pgcName: "H.P. IVA repercutido", grupo: "Pasivo Corriente", subgrupo: "Administraciones Publicas" },
  "26": { pgc: "5530", pgcName: "Socios, cuenta corriente", grupo: "Pasivo Corriente", subgrupo: "Deudas empresas grupo CP" },
  "27": { pgc: "440", pgcName: "Deudores", grupo: "Activo Corriente", subgrupo: "Otros deudores" },
  "41": { pgc: "250", pgcName: "Inversiones financieras a largo plazo", grupo: "Activo No Corriente", subgrupo: "Inversiones financieras LP" },
  "43": { pgc: "213", pgcName: "Maquinaria e instalaciones tecnicas", grupo: "Activo No Corriente", subgrupo: "Inmovilizado material" },
  "51": { pgc: "100", pgcName: "Capital social", grupo: "Patrimonio Neto", subgrupo: "Fondos propios" },
  "55": { pgc: "112", pgcName: "Reserva legal", grupo: "Patrimonio Neto", subgrupo: "Fondos propios" },
  "56": { pgc: "120", pgcName: "Remanente", grupo: "Patrimonio Neto", subgrupo: "Fondos propios" },
  "62": { pgc: "629", pgcName: "Otros servicios", grupo: "Gastos", subgrupo: "Servicios exteriores" },
  "63": { pgc: "640", pgcName: "Sueldos y salarios", grupo: "Gastos", subgrupo: "Gastos de personal" },
  "64": { pgc: "681", pgcName: "Amortizacion del inmovilizado material", grupo: "Gastos", subgrupo: "Amortizaciones" },
  "68": { pgc: "678", pgcName: "Gastos excepcionales", grupo: "Gastos", subgrupo: "Otros resultados" },
  "72": { pgc: "705", pgcName: "Prestaciones de servicios", grupo: "Ingresos", subgrupo: "Importe neto cifra negocios" },
  "81": { pgc: "129", pgcName: "Resultado del ejercicio", grupo: "Patrimonio Neto", subgrupo: "Fondos propios" }
};

const REQUIRED_FIELDS = ["code", "name", "sid", "sia", "cargos", "abonos", "sfd", "sfa"];

const BALANCE_GROUPS = {
  "Activo No Corriente": { items: [], totalMXN: 0, totalEUR: 0 },
  "Activo Corriente": { items: [], totalMXN: 0, totalEUR: 0 },
  "Patrimonio Neto": { items: [], totalMXN: 0, totalEUR: 0 },
  "Pasivo No Corriente": { items: [], totalMXN: 0, totalEUR: 0 },
  "Pasivo Corriente": { items: [], totalMXN: 0, totalEUR: 0 }
};

const PNL_SECTIONS = {
  "Importe neto cifra negocios": { items: [], totalMXN: 0, totalEUR: 0 },
  "Otros ingresos de explotacion": { items: [], totalMXN: 0, totalEUR: 0 },
  "Gastos de personal": { items: [], totalMXN: 0, totalEUR: 0 },
  "Servicios exteriores": { items: [], totalMXN: 0, totalEUR: 0 },
  "Tributos": { items: [], totalMXN: 0, totalEUR: 0 },
  "Amortizaciones": { items: [], totalMXN: 0, totalEUR: 0 },
  "Gastos excepcionales": { items: [], totalMXN: 0, totalEUR: 0 },
  "Resultado financiero": { items: [], totalMXN: 0, totalEUR: 0 },
  "Otros resultados": { items: [], totalMXN: 0, totalEUR: 0 }
};

const normalizeHeader = (value) =>
  String(value ?? "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .replace(/[^a-z0-9]/g, "");

const parseNumber = (value) => {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const clean = String(value ?? "")
    .replace(/\s/g, "")
    .replace(/\./g, "")
    .replace(/,/g, ".")
    .replace(/[^0-9.-]/g, "");
  const parsed = Number.parseFloat(clean);
  return Number.isFinite(parsed) ? parsed : 0;
};

const parseString = (value) => String(value ?? "").trim();
const isZeroSegment = (segment) => /^0+$/.test(String(segment ?? "").trim());

const mapField = (rawKey) => {
  const key = normalizeHeader(rawKey);
  if (["codigo", "codigocuenta", "cuenta", "code", "conta", "cod", "contacontabilistica"].includes(key)) return "code";
  if (["nombre", "nome", "nomedaconta", "descripcion", "descricao", "designacao", "name", "concepto"].includes(key)) return "name";
  if (["sid", "saldoinicialdeudor", "saldoinicialdebe", "saldoinicialdevedor", "saldoinicialdevedora", "debitoacum", "debitoacumulado"].includes(key)) return "sid";
  if (["sia", "saldoinicialacreedor", "saldoinicialhaber", "saldoinicialcredor", "creditoacum", "creditoacumulado"].includes(key)) return "sia";
  if (["cargos", "debe", "movimientodebe", "debito", "debitos", "movimentodebito", "debitomes", "movimentodebitomes"].includes(key)) return "cargos";
  if (["abonos", "haber", "movimientohaber", "credito", "creditos", "movimentocredito", "creditomes", "movimentocreditomes"].includes(key)) return "abonos";
  if (["sfd", "saldofinaldeudor", "saldofinaldebe", "saldofinaldevedor", "saldofinaldevedora", "saldodevedor"].includes(key)) return "sfd";
  if (["sfa", "saldofinalacreedor", "saldofinalhaber", "saldofinalcredor", "saldocredor"].includes(key)) return "sfa";
  return null;
};

export function normalizeRows(rawRows) {
  const rows = rawRows
    .map((row, index) => {
      const transformed = {};
      for (const [k, v] of Object.entries(row)) {
        const mapped = mapField(k);
        if (mapped) transformed[mapped] = v;
      }
      if (!transformed.code) return null;
      return {
        _rowId: parseString(row._rowId || row.rowId || row.id) || `row-${index + 1}`,
        _isNew: Boolean(row._isNew),
        _excludeFromAnalysis: Boolean(row._excludeFromAnalysis),
        code: String(transformed.code).trim(),
        name: String(transformed.name ?? "Sin descripcion").trim(),
        sid: parseNumber(transformed.sid),
        sia: parseNumber(transformed.sia),
        cargos: parseNumber(transformed.cargos),
        abonos: parseNumber(transformed.abonos),
        sfd: parseNumber(transformed.sfd),
        sfa: parseNumber(transformed.sfa)
      };
    })
    .filter(Boolean);

  const missing = REQUIRED_FIELDS.filter((field) => !rows.some((r) => field in r));
  return { rows, missing };
}

const mappingCandidates = (code) => {
  const safeCode = String(code ?? "").trim();
  if (!safeCode) return [];
  const parts = safeCode.split(/[.-]/).filter(Boolean);
  const candidates = [safeCode];
  for (let len = parts.length; len > 0; len -= 1) {
    const prefix = parts.slice(0, len);
    candidates.push(prefix.join("."));
    candidates.push(prefix.join("-"));
  }
  return [...new Set(candidates.filter(Boolean))];
};

const mappingFromPtName = (mapping, code, name) => {
  const text = normalizeHeader(name || "");
  const root = String(code || "").trim().split(".")[0];

  if (root === "24") {
    if (text.includes("pagamentoporconta") || text.includes("retenc")) {
      return { pgc: "473", pgcName: "H.P. retenciones y pagos a cuenta", grupo: "Activo Corriente", subgrupo: "Administraciones Publicas" };
    }
    if (text.includes("iva") && (text.includes("dedut") || text.includes("soport"))) {
      return { pgc: "4720", pgcName: "H.P. IVA soportado", grupo: "Activo Corriente", subgrupo: "Administraciones Publicas" };
    }
    if (text.includes("iva") && (text.includes("liquid") || text.includes("repercut"))) {
      return { pgc: "4770", pgcName: "H.P. IVA repercutido", grupo: "Pasivo Corriente", subgrupo: "Administraciones Publicas" };
    }
  }

  if (root === "68" && text.includes("imposto")) {
    return { pgc: "631", pgcName: "Otros tributos", grupo: "Gastos", subgrupo: "Tributos" };
  }

  if (root === "62") {
    if (text.includes("arrend") || text.includes("aluguer") || text.includes("renda")) {
      return { pgc: "621", pgcName: "Arrendamientos y canones", grupo: "Gastos", subgrupo: "Servicios exteriores" };
    }
    if (text.includes("banc") || text.includes("comissa") || text.includes("comis")) {
      return { pgc: "626", pgcName: "Servicios bancarios y similares", grupo: "Gastos", subgrupo: "Servicios exteriores" };
    }
    if (text.includes("especializ") || text.includes("consult") || text.includes("assessor") || text.includes("profission")) {
      return { pgc: "623", pgcName: "Servicios profesionales independientes", grupo: "Gastos", subgrupo: "Servicios exteriores" };
    }
  }

  return mapping;
};

export function findMapping(code, name = "") {
  const safeCode = String(code ?? "").trim();
  if (!safeCode) return null;
  for (const candidate of mappingCandidates(safeCode)) {
    if (ACCOUNT_MAPPING[candidate]) return ACCOUNT_MAPPING[candidate];
  }
  for (const candidate of mappingCandidates(safeCode)) {
    if (PT_INITIAL_PREFIX_MAPPING[candidate]) {
      return mappingFromPtName(PT_INITIAL_PREFIX_MAPPING[candidate], safeCode, name);
    }
  }
  return null;
}

const accountDisplayValue = (group, saldo) => {
  if (["Pasivo Corriente", "Pasivo No Corriente", "Patrimonio Neto", "Ingresos", "Ingresos Financieros"].includes(group)) {
    return -saldo;
  }
  return saldo;
};

const pnlSignedFromAggregate = (row) => {
  if (row.grupo === "Ingresos" || row.grupo === "Ingresos Financieros") return row.totalMXN;
  if (row.grupo === "Gastos" || row.grupo === "Gastos Financieros") return -row.totalMXN;
  return 0;
};

const pnlSignedFromAggregateEur = (row) => {
  if (row.grupo === "Ingresos" || row.grupo === "Ingresos Financieros") return row.totalEUR;
  if (row.grupo === "Gastos" || row.grupo === "Gastos Financieros") return -row.totalEUR;
  return 0;
};

const normalizeManualMapping = (manual) => {
  if (!manual || typeof manual !== "object") return null;
  const pgc = parseString(manual.pgc || manual.pgcCode);
  const pgcName = parseString(manual.pgcName);
  const grupo = parseString(manual.grupo);
  const subgrupo = parseString(manual.subgrupo);
  if (!pgc && !pgcName && !grupo && !subgrupo) return null;
  const inferredGroup = inferGroupFromPgc(pgc);
  const normalizedGroup = normalizeGroup(grupo);
  const finalGroup = normalizedGroup === "Sin clasificar" ? inferredGroup : normalizedGroup;
  return {
    pgc: pgc || "SIN MAPEO",
    pgcName: pgcName || "Sin equivalencia PGC",
    grupo: finalGroup || "Sin clasificar",
    subgrupo: subgrupo || "Sin clasificar"
  };
};

const normalizeGroup = (groupValue) => {
  const raw = String(groupValue || "").trim();
  if (!raw) return "Sin clasificar";
  const normalized = raw
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .replace(/\s+/g, " ")
    .trim();
  const aliases = {
    "activo corriente": "Activo Corriente",
    "activo no corriente": "Activo No Corriente",
    "pasivo corriente": "Pasivo Corriente",
    "pasivo no corriente": "Pasivo No Corriente",
    "patrimonio neto": "Patrimonio Neto",
    ingresos: "Ingresos",
    "ingresos financieros": "Ingresos Financieros",
    gastos: "Gastos",
    "gastos financieros": "Gastos Financieros",
    "sin clasificar": "Sin clasificar"
  };
  return aliases[normalized] || raw;
};

const inferGroupFromPgc = (pgcCode) => {
  const code = String(pgcCode || "").replace(/\D/g, "");
  if (!code) return "Sin clasificar";

  if (/^(10|11|12|13)/.test(code)) return "Patrimonio Neto";
  if (/^(14|15|16|17|18)/.test(code)) return "Pasivo No Corriente";
  if (/^2/.test(code)) return "Activo No Corriente";
  if (/^3/.test(code)) return "Activo Corriente";

  // Grupo 4 mixto (deudor/acreedor): afinamos por subseries tipicas.
  if (/^(40|41|438|439|47[5-9])/.test(code)) return "Pasivo Corriente";
  if (/^(43|44|46|47[0-4])/.test(code)) return "Activo Corriente";

  // Grupo 5 mixto: 50-52 y 51x deudas CP; 57 tesoreria activo corriente.
  if (/^(50|51|52|56|59)/.test(code)) return "Pasivo Corriente";
  if (/^(53|54|55|57|58)/.test(code)) return "Activo Corriente";

  if (/^6/.test(code)) return "Gastos";
  if (/^7/.test(code)) return "Ingresos";
  if (/^8/.test(code)) return "Gastos";
  if (/^9/.test(code)) return "Ingresos";

  return "Sin clasificar";
};

const detectSummaryLine = (row, allRows) => {
  const code = String(row.code ?? "").trim();
  if (!code) return false;

  // Parent accounts in hierarchical charts (e.g. 21, 21.1, 22.1.1) are
  // summary lines when they have child accounts.
  if (
    allRows.some((candidate) => {
      const candidateCode = String(candidate.code ?? "").trim();
      if (!candidateCode || candidateCode === code) return false;
      return candidateCode.startsWith(`${code}.`) || candidateCode.startsWith(`${code}-`);
    })
  ) {
    return true;
  }

  if (code.startsWith("000-000-")) return true;

  const segments = code.split("-");
  let trailingZeroCount = 0;
  for (let i = segments.length - 1; i >= 0; i -= 1) {
    if (isZeroSegment(segments[i])) trailingZeroCount += 1;
    else break;
  }

  if (trailingZeroCount === 0) return false;
  const prefixLen = segments.length - trailingZeroCount;
  if (prefixLen <= 0) return true;
  const prefix = `${segments.slice(0, prefixLen).join("-")}-`;
  return allRows.some((candidate) => candidate.code !== code && String(candidate.code ?? "").startsWith(prefix));
};

export function convertRows(rows, exchangeRate = 1.0, manualMappings = {}, period = null) {
  const convertedData = rows.map((row) => {
    const manualMapping = normalizeManualMapping(
      manualMappings[row._rowId] || manualMappings[row.code] || row.manualMapping
    );
    const automaticMapping = findMapping(row.code, row.name);
    const mapping = manualMapping || automaticMapping;
    const saldo = row.sfd - row.sfa;
    const saldoEur = saldo * exchangeRate;
    const grupo = mapping?.grupo ?? "Sin clasificar";
    const displayMXN = accountDisplayValue(grupo, saldo);

    const isSummaryLine = detectSummaryLine(row, rows);
    const excludedByUser = Boolean(row._excludeFromAnalysis);
    const excludeFromAnalysis = isSummaryLine || excludedByUser;

    return {
      ...row,
      mapping,
      pgcCode: mapping?.pgc ?? "SIN MAPEO",
      pgcName: mapping?.pgcName ?? "Sin equivalencia PGC",
      grupo,
      subgrupo: mapping?.subgrupo ?? "Sin clasificar",
      saldo,
      saldoEur,
      displayMXN,
      displayEUR: displayMXN * exchangeRate,
      manualMappingApplied: Boolean(manualMapping),
      isSummaryLine,
      excludeFromAnalysis
    };
  });

  const rowsForAnalysis = convertedData.filter((row) => !row.excludeFromAnalysis);

  const aggregateMap = {};
  for (const row of rowsForAnalysis) {
    const key = row.pgcCode;
    if (!aggregateMap[key]) {
      aggregateMap[key] = {
        pgcCode: row.pgcCode,
        pgcName: row.pgcName,
        grupo: row.grupo,
        subgrupo: row.subgrupo,
        totalMXN: 0,
        totalEUR: 0,
        details: []
      };
    }
    aggregateMap[key].totalMXN += row.displayMXN;
    aggregateMap[key].totalEUR += row.displayEUR;
    aggregateMap[key].details.push(row);
  }

  const pgcAggregated = Object.values(aggregateMap).sort((a, b) => a.pgcCode.localeCompare(b.pgcCode));

  const balanceGroups = structuredClone(BALANCE_GROUPS);
  for (const row of pgcAggregated) {
    if (!balanceGroups[row.grupo]) continue;
    balanceGroups[row.grupo].items.push(row);
    balanceGroups[row.grupo].totalMXN += row.totalMXN;
    balanceGroups[row.grupo].totalEUR += row.totalEUR;
  }

  const totalActivoMXN = balanceGroups["Activo No Corriente"].totalMXN + balanceGroups["Activo Corriente"].totalMXN;
  const totalPasivoPNMXN =
    balanceGroups["Patrimonio Neto"].totalMXN +
    balanceGroups["Pasivo No Corriente"].totalMXN +
    balanceGroups["Pasivo Corriente"].totalMXN;
  const totalActivoEUR = balanceGroups["Activo No Corriente"].totalEUR + balanceGroups["Activo Corriente"].totalEUR;
  const totalPasivoPNEUR =
    balanceGroups["Patrimonio Neto"].totalEUR +
    balanceGroups["Pasivo No Corriente"].totalEUR +
    balanceGroups["Pasivo Corriente"].totalEUR;
  const differenceMXN = totalActivoMXN - totalPasivoPNMXN;
  const differenceEUR = totalActivoEUR - totalPasivoPNEUR;

  let adjustedTotalPasivoPNMXN = totalPasivoPNMXN;
  let adjustedTotalPasivoPNEUR = totalPasivoPNEUR;
  let autoResultLine = null;
  if (Math.abs(differenceMXN) > 0.01) {
    autoResultLine = {
      pgcCode: "129",
      pgcName: "Resultado del periodo pendiente de cierre",
      grupo: "Patrimonio Neto",
      subgrupo: "Fondos propios",
      totalMXN: differenceMXN,
      totalEUR: differenceEUR,
      details: []
    };
    adjustedTotalPasivoPNMXN += differenceMXN;
    adjustedTotalPasivoPNEUR += differenceEUR;
  }

  const pnlSections = structuredClone(PNL_SECTIONS);
  for (const row of pgcAggregated) {
    const sectionKey = row.subgrupo === "Gastos financieros" ? "Resultado financiero" : row.subgrupo;
    if (!pnlSections[sectionKey]) continue;
    const signedMXN = pnlSignedFromAggregate(row);
    const signedEUR = pnlSignedFromAggregateEur(row);
    pnlSections[sectionKey].items.push({
      ...row,
      subgrupo: sectionKey,
      totalMXN: signedMXN,
      totalEUR: signedEUR
    });
    pnlSections[sectionKey].totalMXN += signedMXN;
    pnlSections[sectionKey].totalEUR += signedEUR;
  }

  const ingresosMx = pnlSections["Importe neto cifra negocios"].totalMXN + pnlSections["Otros ingresos de explotacion"].totalMXN;
  const gastosMx = -(
    pnlSections["Gastos de personal"].totalMXN +
    pnlSections["Servicios exteriores"].totalMXN +
    pnlSections["Tributos"].totalMXN +
    pnlSections["Amortizaciones"].totalMXN +
    pnlSections["Gastos excepcionales"].totalMXN
  );
  const resultadoFinancieroMx = pnlSections["Resultado financiero"].totalMXN;
  const otrosResultadosMx = pnlSections["Otros resultados"].totalMXN;

  const resultadoExplotacionMx =
    pnlSections["Importe neto cifra negocios"].totalMXN +
    pnlSections["Otros ingresos de explotacion"].totalMXN +
    pnlSections["Gastos de personal"].totalMXN +
    pnlSections["Servicios exteriores"].totalMXN +
    pnlSections["Tributos"].totalMXN +
    pnlSections["Amortizaciones"].totalMXN +
    pnlSections["Gastos excepcionales"].totalMXN;
  const resultadoAntesImpuestosMx = resultadoExplotacionMx + resultadoFinancieroMx + otrosResultadosMx;

  const ingresosEur = pnlSections["Importe neto cifra negocios"].totalEUR + pnlSections["Otros ingresos de explotacion"].totalEUR;
  const gastosEur = -(
    pnlSections["Gastos de personal"].totalEUR +
    pnlSections["Servicios exteriores"].totalEUR +
    pnlSections["Tributos"].totalEUR +
    pnlSections["Amortizaciones"].totalEUR +
    pnlSections["Gastos excepcionales"].totalEUR
  );
  const resultadoFinancieroEur = pnlSections["Resultado financiero"].totalEUR;
  const otrosResultadosEur = pnlSections["Otros resultados"].totalEUR;
  const resultadoExplotacionEur =
    pnlSections["Importe neto cifra negocios"].totalEUR +
    pnlSections["Otros ingresos de explotacion"].totalEUR +
    pnlSections["Gastos de personal"].totalEUR +
    pnlSections["Servicios exteriores"].totalEUR +
    pnlSections["Tributos"].totalEUR +
    pnlSections["Amortizaciones"].totalEUR +
    pnlSections["Gastos excepcionales"].totalEUR;
  const resultadoAntesImpuestosEur = resultadoExplotacionEur + resultadoFinancieroEur + otrosResultadosEur;

  const unmappedRows = rowsForAnalysis.filter((item) => item.pgcCode === "SIN MAPEO");

  const balances = {
    totalDebeInicial: rowsForAnalysis.reduce((acc, r) => acc + r.sid, 0),
    totalHaberInicial: rowsForAnalysis.reduce((acc, r) => acc + r.sia, 0),
    totalDebeFinal: rowsForAnalysis.reduce((acc, r) => acc + r.sfd, 0),
    totalHaberFinal: rowsForAnalysis.reduce((acc, r) => acc + r.sfa, 0)
  };

  return {
    metadata: {
      exchangeRate,
      rowCount: convertedData.length,
      analyzedRowCount: rowsForAnalysis.length,
      summaryExcludedCount: convertedData.filter((row) => row.isSummaryLine).length,
      unmappedCount: unmappedRows.length,
      manualMappingCount: convertedData.filter((row) => row.manualMappingApplied).length,
      mappedCoveragePct: rowsForAnalysis.length ? ((rowsForAnalysis.length - unmappedRows.length) / rowsForAnalysis.length) * 100 : 0,
      period,
      generatedAt: new Date().toISOString()
    },
    convertedData,
    pgcAggregated,
    balanceSheet: {
      groups: balanceGroups,
      totalActivoMXN,
      totalPasivoPNMXN,
      totalActivoEUR,
      totalPasivoPNEUR,
      differenceMXN,
      differenceEUR,
      autoResultLine,
      adjustedTotalPasivoPNMXN,
      adjustedTotalPasivoPNEUR,
      adjustedDifferenceMXN: totalActivoMXN - adjustedTotalPasivoPNMXN,
      adjustedDifferenceEUR: totalActivoEUR - adjustedTotalPasivoPNEUR
    },
    pnl: {
      sections: pnlSections,
      ingresosMx,
      gastosMx,
      resultadoExplotacionMx,
      resultadoFinancieroMx,
      otrosResultadosMx,
      resultadoAntesImpuestosMx,
      ingresosEur,
      gastosEur,
      resultadoExplotacionEur,
      resultadoFinancieroEur,
      otrosResultadosEur,
      resultadoAntesImpuestosEur
    },
    validations: {
      trialBalanceInitialDifference: balances.totalDebeInicial - balances.totalHaberInicial,
      trialBalanceFinalDifference: balances.totalDebeFinal - balances.totalHaberFinal,
      unmappedRows
    }
  };
}

export function parseWorkbookBuffer(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("El archivo no contiene hojas de calculo.");
  }

  const worksheet = workbook.Sheets[firstSheetName];
  const matrixRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

  const headerMaps = matrixRows.map((row = []) => {
    const mapped = {};
    for (let i = 0; i < row.length; i += 1) {
      const field = mapField(row[i]);
      if (field && mapped[field] == null) mapped[field] = i;
    }
    return mapped;
  });
  const headerRowIndex = headerMaps.findIndex((mapped) => mapped.code != null && mapped.name != null);
  const headerMap = headerRowIndex >= 0 ? headerMaps[headerRowIndex] : {};

  if (headerRowIndex >= 0) {
    const cell = (row, field, fallback) => {
      const idx = headerMap[field] ?? fallback;
      return row?.[idx] ?? "";
    };

    let start = headerRowIndex + 1;
    while (start < matrixRows.length) {
      const code = String(cell(matrixRows[start], "code", 0)).trim();
      if (code) break;
      start += 1;
    }

    const rows = [];
    for (let i = start; i < matrixRows.length; i += 1) {
      const row = matrixRows[i] || [];
      const code = String(cell(row, "code", 0)).trim();
      if (!code || !/[0-9]/.test(code)) continue;

      rows.push({
        _rowId: `row-${rows.length + 1}`,
        _isNew: false,
        _excludeFromAnalysis: false,
        code,
        name: String(cell(row, "name", 1)).trim() || "Sin descripcion",
        sid: parseNumber(cell(row, "sid", 2)),
        sia: parseNumber(cell(row, "sia", 3)),
        cargos: parseNumber(cell(row, "cargos", 4)),
        abonos: parseNumber(cell(row, "abonos", 5)),
        sfd: parseNumber(cell(row, "sfd", 6)),
        sfa: parseNumber(cell(row, "sfa", 7))
      });
    }

    if (rows.length > 0) return rows;
  }

  const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
  const { rows } = normalizeRows(rawRows);
  return rows;
}

export function buildExportWorkbook(conversion) {
  const wb = XLSX.utils.book_new();

  const wsMapping = XLSX.utils.json_to_sheet(
    conversion.convertedData.map((r) => ({
      "Cta Portugal": r.code,
      "Nombre Portugal": r.name,
      "Cta PGC": r.pgcCode,
      "Nombre PGC": r.pgcName,
      "Grupo": r.grupo,
      "Subgrupo": r.subgrupo,
      "Saldo base": r.displayMXN,
      "Saldo EUR": r.displayEUR,
      "Linea sumatoria": r.isSummaryLine ? "Si" : "No",
      "Excluida analisis": r.excludeFromAnalysis ? "Si" : "No"
    }))
  );
  XLSX.utils.book_append_sheet(wb, wsMapping, "Mapeo");

  const wsAgg = XLSX.utils.json_to_sheet(
    conversion.pgcAggregated.map((r) => ({
      "Cta PGC": r.pgcCode,
      "Nombre PGC": r.pgcName,
      Grupo: r.grupo,
      Subgrupo: r.subgrupo,
      "Total base": r.totalMXN,
      "Total EUR": r.totalEUR
    }))
  );
  XLSX.utils.book_append_sheet(wb, wsAgg, "Balanza_PGC");

  const wsValidation = XLSX.utils.json_to_sheet([
    {
      Control: "Dif. Balanza Inicial (Debe-Haber)",
      Valor: conversion.validations.trialBalanceInitialDifference
    },
    {
      Control: "Dif. Balanza Final (Debe-Haber)",
      Valor: conversion.validations.trialBalanceFinalDifference
    },
    {
      Control: "Dif. Balance PGC (Activo - PN y Pasivo)",
      Valor: conversion.balanceSheet.differenceMXN
    },
    {
      Control: "Cobertura mapeo (%)",
      Valor: conversion.metadata.mappedCoveragePct
    },
    {
      Control: "Lineas analizadas",
      Valor: conversion.metadata.analyzedRowCount
    },
    {
      Control: "Lineas sumatorias excluidas",
      Valor: conversion.metadata.summaryExcludedCount
    },
    {
      Control: "Sin mapeo",
      Valor: conversion.metadata.unmappedCount
    }
  ]);
  XLSX.utils.book_append_sheet(wb, wsValidation, "Validaciones");

  return wb;
}

export const SAMPLE_ROWS = [
  { code: "101-001-0001", name: "Caja y Efectivo", sid: 82900.95, sia: 0, cargos: 0, abonos: 0, sfd: 82900.95, sfa: 0 },
  { code: "102-001-0001", name: "BBVA Bancomer M.N. 3810", sid: 269952.85, sia: 0, cargos: 1247509.6, abonos: 1405442.13, sfd: 112020.32, sfa: 0 },
  { code: "102-002-0001", name: "BBVA Bancomer USD 6344", sid: 743947.94, sia: 0, cargos: 1228050.96, abonos: 1646992.78, sfd: 325006.12, sfa: 0 },
  { code: "104-001-0001", name: "Duetto Research, Inc.", sid: 21137.07, sia: 0, cargos: 0, abonos: 0, sfd: 21137.07, sfa: 0 },
  { code: "201-001-0000", name: "Proveedores Nacional (varios)", sid: 0, sia: 283254.54, cargos: 259148.68, abonos: 323377.05, sfd: 0, sfa: 347482.91 },
  { code: "203-003-0003", name: "Paraty Hoteles Espana", sid: 0, sia: 4473166.34, cargos: 0, abonos: 88951.14, sfd: 0, sfa: 4562117.48 },
  { code: "208-001-0000", name: "IVA por pagar", sid: 0, sia: 171173.6, cargos: 165656, abonos: 141381.9, sfd: 0, sfa: 146899.5 },
  { code: "301-001-0001", name: "Paraty Hoteles S.L", sid: 0, sia: 49500, cargos: 0, abonos: 0, sfd: 0, sfa: 49500 },
  { code: "401-001-0000", name: "Ventas 16%", sid: 0, sia: 0, cargos: 0, abonos: 1355730.16, sfd: 0, sfa: 1355730.16 },
  { code: "601-000-0001", name: "Sueldos y salarios", sid: 0, sia: 0, cargos: 312154.27, abonos: 0, sfd: 312154.27, sfa: 0 },
  { code: "701-002-0000", name: "Perdida cambiaria", sid: 0, sia: 0, cargos: 120491.28, abonos: 0, sfd: 120491.28, sfa: 0 },
  { code: "702-002-0000", name: "Utilidad cambiaria", sid: 0, sia: 0, cargos: 0, abonos: 7544.04, sfd: 0, sfa: 7544.04 }
];
