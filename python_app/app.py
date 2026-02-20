from __future__ import annotations

import datetime as dt
import uuid
from typing import Any

import pandas as pd
import streamlit as st

from conversion_engine import convert_rows, export_conversion_xlsx, parse_workbook
from db import init_db, list_periods, load_period_data, save_period_data

st.set_page_config(page_title="SNC Portugal a PGC Espana", layout="wide")
init_db()

MONTHS = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}

GROUP_OPTIONS = [
    "Sin clasificar",
    "Activo No Corriente",
    "Activo Corriente",
    "Patrimonio Neto",
    "Pasivo No Corriente",
    "Pasivo Corriente",
    "Ingresos",
    "Ingresos Financieros",
    "Gastos",
    "Gastos Financieros",
]

SUBGROUP_OPTIONS = [
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
    "Otros resultados",
]


st.markdown(
    """
    <style>
      .main h1 { margin-bottom: 0.1rem; }
      .block-container { padding-top: 1.3rem; }
      .status-ok { color: #065f46; font-weight: 700; }
      .status-bad { color: #991b1b; font-weight: 700; }
      .subtle { color: #475569; }
    </style>
    """,
    unsafe_allow_html=True,
)


def fmt(n: float) -> str:
    return f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def ensure_state() -> None:
    now = dt.date.today()
    st.session_state.setdefault("period_month", now.month)
    st.session_state.setdefault("period_year", now.year)
    st.session_state.setdefault("exchange_rate", 1.0)
    st.session_state.setdefault("source_rows", [])
    st.session_state.setdefault("manual_mappings", {})
    st.session_state.setdefault("conversion", None)


def analyze_current() -> None:
    rows = st.session_state["source_rows"]
    mappings = st.session_state["manual_mappings"]
    period = {"month": st.session_state["period_month"], "year": st.session_state["period_year"]}
    st.session_state["conversion"] = convert_rows(rows, st.session_state["exchange_rate"], mappings, period)


def can_save(conversion: dict[str, Any] | None) -> tuple[bool, str]:
    if not conversion:
        return False, "Sin analisis"
    analyzed = conversion["metadata"]["analyzedRowCount"]
    unmapped = conversion["metadata"]["unmappedCount"]
    trial = abs(conversion["validations"]["trialBalanceFinalDifference"])
    if analyzed <= 0:
        return False, "No hay lineas analizadas"
    if unmapped > 0:
        return False, "Hay lineas sin mapear"
    if trial > 0.01:
        return False, "La balanza final no cuadra"
    return True, "Listo para guardar"


def load_period_action(year: int, month: int) -> None:
    payload = load_period_data(year, month)
    if not payload:
        st.warning("No existe informacion guardada para ese periodo")
        return
    st.session_state["source_rows"] = payload["rows"]
    st.session_state["manual_mappings"] = payload["manualMappings"]
    st.session_state["exchange_rate"] = float(payload["period"]["exchange_rate"] or 1.0)
    analyze_current()


def apply_partidas_changes(edited: pd.DataFrame) -> None:
    new_rows = []
    new_maps: dict[str, dict[str, str]] = {}
    for _, row in edited.iterrows():
        row_id = str(row.get("_rowId") or f"row-{uuid.uuid4().hex[:8]}")
        code = str(row.get("code", "")).strip()
        if not code:
            continue

        new_rows.append(
            {
                "_rowId": row_id,
                "_isNew": True,
                "_excludeFromAnalysis": False,
                "code": code,
                "name": str(row.get("name", "")).strip(),
                "sid": float(row.get("sid", 0) or 0),
                "sia": float(row.get("sia", 0) or 0),
                "cargos": float(row.get("cargos", 0) or 0),
                "abonos": float(row.get("abonos", 0) or 0),
                "sfd": float(row.get("sfd", 0) or 0),
                "sfa": float(row.get("sfa", 0) or 0),
            }
        )

        pgc = str(row.get("manual_pgc", "")).strip()
        pgc_name = str(row.get("manual_pgcName", "")).strip()
        grupo = str(row.get("manual_grupo", "Sin clasificar"))
        subgrupo = str(row.get("manual_subgrupo", "Sin clasificar"))
        if pgc or pgc_name or grupo != "Sin clasificar" or subgrupo != "Sin clasificar":
            new_maps[row_id] = {
                "pgc": pgc,
                "pgcName": pgc_name,
                "grupo": grupo,
                "subgrupo": subgrupo,
            }

    st.session_state["source_rows"] = new_rows
    st.session_state["manual_mappings"] = new_maps
    analyze_current()


ensure_state()

st.title("SNC Portugal a PGC Espana")
st.caption("Flujo recomendado: 1) Subir archivo 2) Revisar y mapear 3) Guardar periodo")

with st.sidebar:
    st.subheader("Periodo")
    st.session_state["period_month"] = st.selectbox(
        "Mes",
        options=list(MONTHS.keys()),
        format_func=lambda x: MONTHS[x],
        index=list(MONTHS.keys()).index(st.session_state["period_month"]),
    )
    st.session_state["period_year"] = st.number_input(
        "Año",
        min_value=2000,
        max_value=2100,
        value=int(st.session_state["period_year"]),
    )
    st.session_state["exchange_rate"] = st.number_input(
        "TC EUR/EUR",
        min_value=0.0001,
        value=float(st.session_state["exchange_rate"]),
        format="%.4f",
    )

    if st.button("Cargar periodo", use_container_width=True):
        load_period_action(int(st.session_state["period_year"]), int(st.session_state["period_month"]))

    upload = st.file_uploader("Subir y analizar archivo", type=["xlsx", "xls", "csv"])
    if upload is not None:
        rows = parse_workbook(upload.read(), upload.name)
        st.session_state["source_rows"] = rows
        st.session_state["manual_mappings"] = {}
        analyze_current()
        st.success(f"Archivo analizado: {len(rows)} lineas")

    if st.button("Recalcular", use_container_width=True):
        analyze_current()

    st.divider()
    st.subheader("Periodos guardados")
    periods = list_periods()
    if periods:
        labels = [f"{p['month']:02d}/{p['year']} · {p['row_count']} lineas" for p in periods]
        idx = st.selectbox("Selecciona", options=range(len(labels)), format_func=lambda i: labels[i], label_visibility="collapsed")
        if st.button("Cargar seleccionado", use_container_width=True):
            p = periods[idx]
            st.session_state["period_month"] = int(p["month"])
            st.session_state["period_year"] = int(p["year"])
            load_period_action(int(p["year"]), int(p["month"]))
    else:
        st.caption("Sin periodos guardados")

conversion = st.session_state["conversion"]
if not conversion:
    st.info("Carga un archivo o selecciona un periodo guardado para empezar")
    st.stop()

ok_save, save_msg = can_save(conversion)

m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("Lineas archivo", conversion["metadata"]["rowCount"])
m2.metric("Lineas analizadas", conversion["metadata"]["analyzedRowCount"])
m3.metric("Sumatorias excluidas", conversion["metadata"]["summaryExcludedCount"])
m4.metric("Sin mapear", conversion["metadata"]["unmappedCount"])
m5.metric("Cobertura mapeo", f"{conversion['metadata']['mappedCoveragePct']:.2f}%")
m6.metric("Total PyG", fmt(conversion["pnl"]["resultadoAntesImpuestosMx"]))

status_cls = "status-ok" if ok_save else "status-bad"
st.markdown(
    f"<div class='{status_cls}'>Estado guardado: {save_msg}</div><div class='subtle'>Dif. balanza final: {fmt(conversion['validations']['trialBalanceFinalDifference'])}</div>",
    unsafe_allow_html=True,
)

act1, act2 = st.columns([1, 1])
with act1:
    if st.button("Guardar periodo en BBDD", disabled=not ok_save, use_container_width=True):
        save_period_data(
            year=int(st.session_state["period_year"]),
            month=int(st.session_state["period_month"]),
            filename="manual-save",
            exchange_rate=float(st.session_state["exchange_rate"]),
            rows=st.session_state["source_rows"],
            manual_mappings=st.session_state["manual_mappings"],
            uploaded_at=dt.datetime.now().isoformat(),
        )
        st.success("Periodo guardado")

with act2:
    xbytes = export_conversion_xlsx(conversion)
    st.download_button(
        "Exportar XLSX",
        data=xbytes,
        file_name=f"conversion_pgc_{st.session_state['period_year']}-{str(st.session_state['period_month']).zfill(2)}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

tabs = st.tabs(["Partidas", "Mapeo", "Balance", "P&G", "Control"])

with tabs[0]:
    f1, f2 = st.columns([1.2, 2])
    with f1:
        status_filter = st.selectbox("Estado", ["todos", "sin-mapear", "mapeadas", "sumatorias"])
    with f2:
        detail_search = st.text_input("Buscar por cuenta o descripcion")

    converted_by_id = {r["_rowId"]: r for r in conversion["convertedData"]}
    display_rows: list[dict[str, Any]] = []
    for r in st.session_state["source_rows"]:
        c = converted_by_id.get(r["_rowId"], {})
        is_summary = bool(c.get("isSummaryLine"))
        is_mapped = c.get("pgcCode") not in (None, "SIN MAPEO")
        status = "sumatoria" if is_summary else ("mapeada" if is_mapped else "sin-mapear")

        if status_filter == "sin-mapear" and status != "sin-mapear":
            continue
        if status_filter == "mapeadas" and status != "mapeada":
            continue
        if status_filter == "sumatorias" and status != "sumatoria":
            continue
        if detail_search:
            q = detail_search.lower()
            if q not in r["code"].lower() and q not in r["name"].lower():
                continue

        manual = st.session_state["manual_mappings"].get(r["_rowId"], {})
        display_rows.append(
            {
                "_rowId": r["_rowId"],
                "code": r["code"],
                "name": r["name"],
                "sid": r["sid"],
                "sia": r["sia"],
                "cargos": r["cargos"],
                "abonos": r["abonos"],
                "sfd": r["sfd"],
                "sfa": r["sfa"],
                "suma_debe": r["sid"] + r["cargos"],
                "suma_haber": r["sia"] + r["abonos"],
                "saldo_neto": r["sfd"] - r["sfa"],
                "estado": status,
                "pgc_asignado": c.get("pgcCode", "SIN MAPEO"),
                "nombre_asignado": c.get("pgcName", "Sin equivalencia PGC"),
                "manual_pgc": manual.get("pgc", ""),
                "manual_pgcName": manual.get("pgcName", ""),
                "manual_grupo": manual.get("grupo", "Sin clasificar"),
                "manual_subgrupo": manual.get("subgrupo", "Sin clasificar"),
            }
        )

    df = pd.DataFrame(display_rows)
    if df.empty:
        st.warning("Sin lineas para mostrar")
    else:
        edited = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "_rowId": st.column_config.TextColumn("rowId", disabled=True),
                "suma_debe": st.column_config.NumberColumn("Suma Debe", disabled=True),
                "suma_haber": st.column_config.NumberColumn("Suma Haber", disabled=True),
                "saldo_neto": st.column_config.NumberColumn("Saldo Neto", disabled=True),
                "estado": st.column_config.TextColumn("Estado", disabled=True),
                "pgc_asignado": st.column_config.TextColumn("PGC asignado", disabled=True),
                "nombre_asignado": st.column_config.TextColumn("Nombre asignado", disabled=True),
                "manual_grupo": st.column_config.SelectboxColumn("Manual Grupo", options=GROUP_OPTIONS),
                "manual_subgrupo": st.column_config.SelectboxColumn("Manual Subgrupo", options=SUBGROUP_OPTIONS),
            },
        )

        if st.button("Aplicar cambios de partidas"):
            apply_partidas_changes(edited)
            st.success("Cambios aplicados")

with tabs[1]:
    st.dataframe(pd.DataFrame(conversion["pgcAggregated"]), use_container_width=True)

with tabs[2]:
    b = conversion["balanceSheet"]
    left, right = st.columns(2)
    with left:
        st.subheader("Activo")
        st.write(f"Total Activo: {fmt(b['totalActivoMXN'])}")
        st.dataframe(
            pd.DataFrame(
                [
                    {"Grupo": k, "Total MXN": v["totalMXN"]}
                    for k, v in b["groups"].items()
                    if "Activo" in k
                ]
            ),
            use_container_width=True,
            hide_index=True,
        )
    with right:
        st.subheader("Patrimonio Neto y Pasivo")
        st.write(f"Total PN + Pasivo: {fmt(b['totalPasivoPNMXN'])}")
        if b.get("autoResultLine"):
            st.warning(f"129 Resultado periodo (ajuste tecnico): {fmt(b['autoResultLine']['totalMXN'])}")
            st.write(f"Total PN + Pasivo (ajustado): {fmt(b['adjustedTotalPasivoPNMXN'])}")
        st.dataframe(
            pd.DataFrame(
                [
                    {"Grupo": k, "Total MXN": v["totalMXN"]}
                    for k, v in b["groups"].items()
                    if "Pasivo" in k or "Patrimonio" in k
                ]
            ),
            use_container_width=True,
            hide_index=True,
        )

with tabs[3]:
    st.subheader("Cuenta de Perdidas y Ganancias")
    st.success(f"Total del periodo (Resultado PyG): {fmt(conversion['pnl']['resultadoAntesImpuestosMx'])}")

    col_a, col_b = st.columns([1, 2])
    with col_a:
        collapse_all = st.toggle("Colapsar secciones", value=False)

    for section, content in conversion["pnl"]["sections"].items():
        with st.expander(f"{section} · Subtotal {fmt(content['totalMXN'])}", expanded=not collapse_all):
            for item in content["items"]:
                st.write(f"{item['pgcCode']} - {item['pgcName']}: {fmt(item['totalMXN'])}")

    st.info(f"Total del periodo: {fmt(conversion['pnl']['resultadoAntesImpuestosMx'])}")

with tabs[4]:
    st.subheader("Control")
    st.write(f"Dif. balanza inicial: {fmt(conversion['validations']['trialBalanceInitialDifference'])}")
    st.write(f"Dif. balanza final: {fmt(conversion['validations']['trialBalanceFinalDifference'])}")
    st.write(f"Sin mapear: {conversion['metadata']['unmappedCount']}")

    if conversion["validations"]["unmappedRows"]:
        st.markdown("**Lineas sin mapear**")
        st.dataframe(pd.DataFrame(conversion["validations"]["unmappedRows"]), use_container_width=True)
