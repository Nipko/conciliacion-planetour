"""
Panel de Control Ejecutivo y Tablero de Conciliación Bancaria - Planetour SAS.
Interfaz moderna, intuitiva, con navegación interactiva y descarga individual de Excel por sección.
"""

import streamlit as st
import pandas as pd
import os
import io
import re
from datetime import datetime

from src.engine.matcher import ReconciliationEngine
from src.reports.excel_generator import ExcelReportGenerator
from src.engine.account_map import ACCOUNTS_CATALOG

# Configuración de página
st.set_page_config(
    page_title="Conciliación Bancaria - Planetour SAS",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS profesionales
st.markdown("""
<style>
    .main-header-title {
        font-size: 26px;
        font-weight: 800;
        color: #1B365D;
        margin-bottom: 2px;
    }
    .main-header-sub {
        font-size: 14px;
        color: #475569;
        margin-bottom: 18px;
    }
    .action-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px 18px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-bottom: 12px;
    }
    .action-card-title {
        font-size: 13px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .action-card-val {
        font-size: 24px;
        font-weight: 800;
        color: #0f172a;
        margin: 4px 0;
    }
    .action-card-desc {
        font-size: 12.5px;
        color: #64748b;
        line-height: 1.4;
    }
    .info-banner {
        background-color: #f8fafc;
        border-left: 5px solid #3b82f6;
        padding: 14px 18px;
        border-radius: 6px;
        margin-bottom: 20px;
        font-size: 13.5px;
        color: #1e293b;
    }
    .danger-banner {
        background-color: #fff1f2;
        border-left: 5px solid #f43f5e;
        padding: 14px 18px;
        border-radius: 6px;
        margin-bottom: 20px;
        font-size: 13.5px;
        color: #881337;
    }
    .warning-banner {
        background-color: #fffbeb;
        border-left: 5px solid #f59e0b;
        padding: 14px 18px;
        border-radius: 6px;
        margin-bottom: 20px;
        font-size: 13.5px;
        color: #78350f;
    }
    .badge-danger {
        background-color: #fee2e2;
        color: #b91c1c;
        padding: 3px 8px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 11px;
    }
    .badge-warning {
        background-color: #fef3c7;
        color: #b45309;
        padding: 3px 8px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 11px;
    }
    .badge-success {
        background-color: #dcfce7;
        color: #15803d;
        padding: 3px 8px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 11px;
    }
    /* Ocultar barra flotante de Streamlit para no generar confusión con descargas CSV */
    [data-testid="stElementToolbar"],
    button[title="Download as CSV"],
    div[data-testid="stDataFrameToolbar"] {
        display: none !important;
    }
</style>
""", unsafe_allow_html=True)

ROOT_DIR = "."
engine = ReconciliationEngine(root_dir=ROOT_DIR)
generator = ExcelReportGenerator()

def safe_export_dataframe_to_excel_bytes(df: pd.DataFrame, title: str, sheet_name: str = "Datos") -> bytes:
    """Exporta DataFrame a Excel con fallback seguro garantizado."""
    if hasattr(generator, "export_dataframe_to_excel_bytes"):
        try:
            return generator.export_dataframe_to_excel_bytes(df, title, sheet_name)
        except Exception:
            pass
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return buf.getvalue()

# Mapeo global de nombres legibles para cuentas bancarias
GLOBAL_ACCOUNT_LABELS = {}
for cfg in ACCOUNTS_CATALOG.values():
    last_digits = str(cfg.account_number)[-4:] if cfg.account_number else ""
    suffix = f" (#{last_digits})" if last_digits else ""
    GLOBAL_ACCOUNT_LABELS[cfg.karing_name] = f"{cfg.bank_name} - {cfg.karing_name}{suffix}"

account_labels = GLOBAL_ACCOUNT_LABELS

def get_account_label(acc_name: str) -> str:
    if acc_name == "TODAS LAS CUENTAS":
        return "🏦 TODAS LAS CUENTAS (Consolidado)"
    return GLOBAL_ACCOUNT_LABELS.get(acc_name, acc_name)

# -------------------------------------------------------------
# GESTIÓN DE ESTADO DE NAVEGACIÓN (DRILL-DOWN INTERACTIVO)
# -------------------------------------------------------------
MENU_OPTIONS = [
    "📊 1. Tablero Ejecutivo de Control",
    "🚨 2. Lo que está MAL (Cruces Intercuentas)",
    "❓ 3. Lo que FALTA por Conciliar (Pendientes)",
    "💸 4. Gastos e Impuestos Bancarios (4x1000)",
    "💰 5. Intereses y Rendimientos de Cuentas",
    "✅ 6. Lo que está BIEN (Conciliados 1 a 1)",
    "🔍 7. Buscador y Auditoría de Montos Repetidos",
    "📤 8. Cargar Nuevos Archivos / Meses"
]

if "active_tab" not in st.session_state or st.session_state.get("active_tab") not in MENU_OPTIONS:
    st.session_state["active_tab"] = MENU_OPTIONS[0]

def navigate_to(tab_name: str):
    st.session_state["active_tab"] = tab_name
    st.rerun()

# -------------------------------------------------------------
# BARRA LATERAL (SIDEBAR): MENÚ PRINCIPAL
# -------------------------------------------------------------
st.sidebar.markdown("## ✈️ PLANETOUR S.A.S.")
st.sidebar.markdown("**Sistema de Conciliación Bancaria**")
st.sidebar.markdown("---")

st.sidebar.markdown("### 🧭 Menú de Navegación")
current_index = MENU_OPTIONS.index(st.session_state["active_tab"]) if st.session_state["active_tab"] in MENU_OPTIONS else 0
selected_menu = st.sidebar.radio(
    "Ir a la sección:",
    options=MENU_OPTIONS,
    index=current_index,
    label_visibility="collapsed"
)

# Sincronizar navegación
if selected_menu != st.session_state["active_tab"]:
    st.session_state["active_tab"] = selected_menu
    st.rerun()

st.sidebar.markdown("---")

# -------------------------------------------------------------
# ENCABEZADO SUPERIOR PRINCIPAL (MES Y CONTROLES VISIBLES Y CLAROS)
# -------------------------------------------------------------
available_months = engine.scan_available_months()
if not available_months:
    available_months = ["AGOSTO 2026", "JULIO 2026", "JUNIO 2026"]

top_col_title, top_col_month, top_col_tol, top_col_btn = st.columns([2.6, 1.4, 0.9, 1.3])

with top_col_title:
    st.markdown('<div class="main-header-title">Tablero de Conciliación Bancaria</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-header-sub">Auditoría contable automática entre Extractos Bancarios y Libros Karing</div>', unsafe_allow_html=True)

with top_col_month:
    selected_month = st.selectbox(
        "📅 Mes a Auditar:",
        options=available_months,
        index=len(available_months) - 1,
        help="Selecciona el mes a auditar. Al cambiar de mes, el sistema procesará de inmediato los extractos y auxiliares."
    )

with top_col_tol:
    tolerance_days = st.number_input(
        "⏱️ Tolerancia:",
        min_value=1,
        max_value=10,
        value=4,
        step=1,
        help="Margen de días entre la fecha en Karing y la fecha en el extracto."
    )

with top_col_btn:
    st.write("")
    st.write("")
    if st.button("🔄 Recalcular / Refrescar", width="stretch", help="Limpia la memoria temporal y reanaliza todos los extractos y libros auxiliares actualizados del disco."):
        st.cache_data.clear()
        st.rerun()

st.markdown("---")

# -------------------------------------------------------------
# EJECUCIÓN DEL MOTOR CON INDICADOR DE CARGA VISIBLE
# -------------------------------------------------------------
@st.cache_data(show_spinner=False)
def run_reconciliation(month: str, tol: int, _cache_version: str = "v2.5"):
    eng = ReconciliationEngine(root_dir=ROOT_DIR, date_tolerance_days=tol)
    return eng.reconcile_month(month)

# Indicador de carga grande y visible al cambiar de mes
with st.spinner(f"⏳ Auditando y procesando extractos y libros contables de {selected_month}... Por favor espere unos segundos."):
    reconciliation = run_reconciliation(selected_month, tolerance_days)

glob = reconciliation["resumen_global"]

# Descarga del libro maestro completo en sidebar
master_excel_name = f"Conciliacion_Planetour_{selected_month.replace(' ', '_')}_MAESTRO.xlsx"
master_excel_data = generator.generate_report_bytes(reconciliation)

st.sidebar.markdown("### 📦 Libro Maestro Consolidado")
st.sidebar.download_button(
    label="📊 Descargar Libro Maestro (.xlsx)",
    data=master_excel_data,
    file_name=master_excel_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    width="stretch",
    help="Descarga el archivo Excel .xlsx oficial consolidado con todas las 6 hojas y fórmulas."
)

# Alertas de Periodo o Inconsistencias de Auditoría
if reconciliation["advertencias"]:
    for adv in reconciliation["advertencias"]:
        st.warning(f"⚠️ **Alerta de Inconsistencia:** {adv}")

# =============================================================
# VISTA 1: TABLERO EJECUTIVO DE CONTROL
# =============================================================
if st.session_state["active_tab"] == MENU_OPTIONS[0]:
    st.markdown("""
    <div class="info-banner">
        <b>¿Cómo interpretar este tablero?</b> Cada tarjeta representa una categoría clave de auditoría. 
        Haz clic en el botón de cada tarjeta para <b>revisar el detalle específico</b> de esa sección o exportarlo a Excel.
    </div>
    """, unsafe_allow_html=True)

    # Barra de Progreso y Semáforo Global
    tasa = glob.get("tasa_conciliacion", 0.0)
    prog_col1, prog_col2 = st.columns([3, 1])
    with prog_col1:
        st.markdown(f"**Tasa de Conciliación de Partidas Bancarias: {tasa}%**")
        st.progress(min(tasa / 100.0, 1.0))
    with prog_col2:
        if tasa >= 80:
            st.markdown('<div style="text-align:center; padding-top:18px;"><span class="badge-success">🟢 Conciliación Muy Avanzada</span></div>', unsafe_allow_html=True)
        elif tasa >= 60:
            st.markdown('<div style="text-align:center; padding-top:18px;"><span class="badge-warning">🟡 Requiere Reclasificación</span></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div style="text-align:center; padding-top:18px;"><span class="badge-danger">🔴 Partidas Pendientes Críticas</span></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # 5 TARJETAS ACCIONABLES CON BOTÓN CLICKABLE
    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        st.markdown(f"""
        <div class="action-card" style="border-top: 4px solid #10b981;">
            <div class="action-card-title" style="color: #059669;">✅ Conciliados (1 a 1)</div>
            <div class="action-card-val">{glob['total_conciliados']:,} <span style="font-size:14px; color:#64748b;">partidas</span></div>
            <div class="action-card-desc">
                Monto: <b>${glob['monto_total_conciliado']:,.2f}</b><br>
                Coincidencia exacta de valor y fecha entre banco y Karing.
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("👉 Ver Conciliados", key="btn_go_conc", width="stretch"):
            navigate_to(MENU_OPTIONS[5])

    with c2:
        st.markdown(f"""
        <div class="action-card" style="border-top: 4px solid #f59e0b;">
            <div class="action-card-title" style="color: #d97706;">🚨 Lo que está MAL (Intercuentas)</div>
            <div class="action-card-val">{glob['total_cruces_intercuentas']:,} <span style="font-size:14px; color:#64748b;">traslados</span></div>
            <div class="action-card-desc">
                Monto: <b>${glob['monto_total_cruces']:,.2f}</b><br>
                Dinero en banco asentado en otra cuenta en Karing.
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("👉 Corregir Cuentas", key="btn_go_cross", width="stretch"):
            navigate_to(MENU_OPTIONS[1])

    with c3:
        st.markdown(f"""
        <div class="action-card" style="border-top: 4px solid #8b5cf6;">
            <div class="action-card-title" style="color: #7c3aed;">💸 Gastos y 4x1000</div>
            <div class="action-card-val">${glob['total_gastos_bancarios']:,.0f}</div>
            <div class="action-card-desc">
                GMF: <b>${glob['total_gmf']:,.0f}</b> | Comis: <b>${glob['total_comisiones']:,.0f}</b><br>
                Cobros bancarios por contabilizar en Karing.
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("👉 Ver Gastos y 4x1000", key="btn_go_tax", width="stretch"):
            navigate_to(MENU_OPTIONS[3])

    with c4:
        st.markdown(f"""
        <div class="action-card" style="border-top: 4px solid #0284c7;">
            <div class="action-card-title" style="color: #0369a1;">💰 Intereses / Rendimientos</div>
            <div class="action-card-val">${glob['total_intereses']:,.0f}</div>
            <div class="action-card-desc">
                Abonos de intereses en ahorro y CDT.<br>
                Ingresos a favor por causar en Karing (421005).
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("👉 Ver Intereses (421005)", key="btn_go_interest", width="stretch"):
            navigate_to(MENU_OPTIONS[4])

    with c5:
        total_faltantes = glob['total_pendientes_banco'] + glob['total_pendientes_karing']
        st.markdown(f"""
        <div class="action-card" style="border-top: 4px solid #ef4444;">
            <div class="action-card-title" style="color: #dc2626;">❓ Lo que FALTA por Conciliar</div>
            <div class="action-card-val">{total_faltantes:,} <span style="font-size:14px; color:#64748b;">partidas</span></div>
            <div class="action-card-desc">
                Banco sin Karing: <b>{glob['total_pendientes_banco']:,}</b><br>
                Karing sin Banco: <b>{glob['total_pendientes_karing']:,}</b>
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("👉 Ver Partidas Faltantes", key="btn_go_pend", width="stretch"):
            navigate_to(MENU_OPTIONS[2])

    st.markdown("<br>", unsafe_allow_html=True)

    # Resumen Consolidado de Saldos
    st.markdown("### ⚖️ Cuadre Consolidado de Saldos")
    b_col1, b_col2, b_col3 = st.columns(3)
    with b_col1:
        st.metric("🏦 Saldo Total en Extractos Bancarios", f"${glob['saldo_total_bancos']:,.2f}")
    with b_col2:
        st.metric("📖 Saldo Total en Libros Karing", f"${glob['saldo_total_karing']:,.2f}")
    with b_col3:
        st.metric(
            "Diferencia Global a Explicar",
            f"${glob['diferencia_global']:,.2f}",
            delta=f"-${glob['total_gastos_bancarios']:,.2f} en gastos banco" if glob['total_gastos_bancarios'] > 0 else None,
            delta_color="off"
        )

    st.markdown("---")

    # Tabla Ejecutiva por Banco y Cuenta
    st.markdown("### 📋 Resumen por Entidad Financiera y Cuenta Contable")
    
    rows_summary = []
    for acc_k, acc_r in reconciliation["cuentas"].items():
        cfg = acc_r["config"]
        m = acc_r["metrics"]

        if m["saldo_banco"] == 0 and m["saldo_karing"] == 0 and m["total_conciliados"] == 0:
            estado_lbl = "⚪ Sin Movimiento"
        elif m["num_pendientes_banco"] == 0 and m["num_pendientes_karing"] == 0:
            estado_lbl = "🟢 Cuadrada"
        elif m["num_cruces_intercuentas"] > 0 and m["num_pendientes_banco"] <= 5:
            estado_lbl = "🟡 Reclasificar"
        else:
            estado_lbl = "🔴 Faltantes"

        rows_summary.append({
            "Estado": estado_lbl,
            "Banco": cfg.bank_name,
            "Cuenta Contable": cfg.karing_name,
            "Saldo Banco ($)": m["saldo_banco"],
            "Saldo Karing ($)": m["saldo_karing"],
            "Diferencia ($)": m["diferencia_bruta"],
            "Conciliados ($)": m["monto_conciliado"],
            "Conc. (#)": m["total_conciliados"],
            "Cruces (#)": m["num_cruces_intercuentas"],
            "Falta Karing (#)": m["num_pendientes_banco"],
            "Falta Banco (#)": m["num_pendientes_karing"],
            "Gastos ($)": m["total_gastos_bancarios"],
            "Cód. Karing": cfg.karing_code,
            "No. Cuenta": str(cfg.account_number)
        })

    cols_order_summary = [
        "Estado", "Banco", "Cuenta Contable",
        "Saldo Banco ($)", "Saldo Karing ($)", "Diferencia ($)",
        "Conciliados ($)", "Conc. (#)", "Cruces (#)",
        "Falta Karing (#)", "Falta Banco (#)", "Gastos ($)",
        "Cód. Karing", "No. Cuenta"
    ]
    df_summary_page = pd.DataFrame(rows_summary)[cols_order_summary]

    # Botón de descarga de esta tabla
    excel_resumen_bytes = safe_export_dataframe_to_excel_bytes(
        df_summary_page,
        title=f"Resumen de Conciliación por Cuentas - {selected_month}",
        sheet_name="Resumen_Cuentas"
    )
    col_dl1, col_dl2 = st.columns([3, 1])
    with col_dl2:
        st.download_button(
            label="📥 Descargar este Resumen en Excel (.xlsx)",
            data=excel_resumen_bytes,
            file_name=f"Resumen_Cuentas_{selected_month.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch"
        )

    st.dataframe(
        df_summary_page,
        column_config={
            "Estado": st.column_config.TextColumn("Estado", width=140),
            "Banco": st.column_config.TextColumn("Banco", width=120),
            "Cuenta Contable": st.column_config.TextColumn("Cuenta Contable", width=220),
            "Saldo Banco ($)": st.column_config.NumberColumn("Saldo Banco", format="$ %,.2f", width=130),
            "Saldo Karing ($)": st.column_config.NumberColumn("Saldo Karing", format="$ %,.2f", width=130),
            "Diferencia ($)": st.column_config.NumberColumn("Diferencia", format="$ %,.2f", width=130),
            "Conciliados ($)": st.column_config.NumberColumn("Conciliados ($)", format="$ %,.2f", width=130),
            "Conc. (#)": st.column_config.NumberColumn("Conc. (#)", format="%d", width=85),
            "Cruces (#)": st.column_config.NumberColumn("Cruces (#)", format="%d", width=85),
            "Falta Karing (#)": st.column_config.NumberColumn("Falta Karing (#)", format="%d", width=110),
            "Falta Banco (#)": st.column_config.NumberColumn("Falta Banco (#)", format="%d", width=105),
            "Gastos ($)": st.column_config.NumberColumn("Gastos ($)", format="$ %,.2f", width=110),
            "Cód. Karing": st.column_config.NumberColumn("Cód. Karing", format="%d", width=95),
            "No. Cuenta": st.column_config.TextColumn("No. Cuenta", width=130)
        },
        width="stretch",
        hide_index=True
    )

# =============================================================
# VISTA 2: LO QUE ESTÁ MAL (CRUCES INTERCUENTAS)
# =============================================================
elif st.session_state["active_tab"] == MENU_OPTIONS[1]:
    st.markdown("""
    <div class="danger-banner">
        <h4 style="margin:0 0 6px 0; color:#881337;">🚨 ¿Qué es esta sección? (Lo que está MAL y requiere corrección)</h4>
        <b>Error detectado:</b> El dinero llegó efectivamente al extracto bancario de la cuenta <b>A</b>, 
        pero en Karing el movimiento fue registrado por error en el auxiliar de la cuenta <b>B</b>.<br>
        <b>Impacto:</b> La cuenta A queda con dinero sobrante sin justificar, y la cuenta B queda con recibos inflados sin respaldo bancario.<br>
        <b>Solución:</b> Realizar el traslado contable en Karing del documento indicado.
    </div>
    """, unsafe_allow_html=True)

    all_cross = []
    for acc_k, acc_r in reconciliation["cuentas"].items():
        all_cross.extend(acc_r.get("cruces_intercuentas", []))

    if all_cross:
        df_cross_full = pd.DataFrame(all_cross)

        m1, m2 = st.columns(2)
        with m1:
            st.metric("Total de Movimientos en Cuenta Errónea", f"{len(df_cross_full):,} casos")
        with m2:
            st.metric("Monto Total a Reclasificar en Karing", f"${df_cross_full['valor'].sum():,.2f}")

        # Filtros de búsqueda para auditoría
        st.markdown("##### Filtros Rápidos:")
        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1:
            bancos_reales = ["TODOS"] + sorted(list(df_cross_full["banco_real"].unique()))
            sel_banco_real = st.selectbox("Filtrar por Banco donde entró el dinero:", bancos_reales)
        with f_col2:
            cuentas_erroneas = ["TODAS"] + sorted(list(df_cross_full["cuenta_registrada_karing"].unique()))
            sel_cta_err = st.selectbox("Filtrar por Cuenta donde se registró erróneamente:", cuentas_erroneas)
        with f_col3:
            search_doc = st.text_input("Buscar por No. de Recibo / Documento:")

        df_show = df_cross_full.copy()
        if sel_banco_real != "TODOS":
            df_show = df_show[df_show["banco_real"] == sel_banco_real]
        if sel_cta_err != "TODAS":
            df_show = df_show[df_show["cuenta_registrada_karing"] == sel_cta_err]
        if search_doc.strip():
            df_show = df_show[df_show["documento_karing"].astype(str).str.contains(search_doc.strip(), case=False, na=False)]

        df_display = df_show[[
            "banco_real", "valor", "cuenta_registrada_karing", "documento_karing",
            "fecha_banco", "fecha_karing", "accion_recomendada", "detalle_karing", "descripcion_banco"
        ]].rename(columns={
            "banco_real": "Banco Real (Entró)",
            "valor": "Valor ($)",
            "cuenta_registrada_karing": "Cuenta Errónea en Karing",
            "documento_karing": "Recibo / Egreso #",
            "fecha_banco": "Fecha Extracto",
            "fecha_karing": "Fecha Karing",
            "accion_recomendada": "Instrucción de Reclasificación",
            "detalle_karing": "Detalle en Karing",
            "descripcion_banco": "Detalle en Extracto"
        })

        # Botón de Descarga individual en Excel
        excel_cross_bytes = safe_export_dataframe_to_excel_bytes(
            df_display,
            title=f"Cruces Intercuentas y Reclasificaciones - {selected_month}",
            sheet_name="Cruces_Intercuentas"
        )
        d_col1, d_col2 = st.columns([3, 1])
        with d_col2:
            st.download_button(
                label="📥 Descargar Listado de Cruces en Excel (.xlsx)",
                data=excel_cross_bytes,
                file_name=f"Cruces_Intercuentas_{selected_month.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch"
            )

        st.dataframe(
            df_display,
            column_config={
                "Banco Real (Entró)": st.column_config.TextColumn("Banco Real", width=140),
                "Valor ($)": st.column_config.NumberColumn("Valor ($)", format="$ %,.2f", width=130),
                "Cuenta Errónea en Karing": st.column_config.TextColumn("Cuenta Errónea", width=170),
                "Recibo / Egreso #": st.column_config.TextColumn("Doc #", width=105),
                "Fecha Extracto": st.column_config.TextColumn("Fecha Banco", width=105),
                "Fecha Karing": st.column_config.TextColumn("Fecha Karing", width=105),
                "Instrucción de Reclasificación": st.column_config.TextColumn("Instrucción Contable", width=380),
                "Detalle en Karing": st.column_config.TextColumn("Detalle Karing", width=250),
                "Detalle en Extracto": st.column_config.TextColumn("Detalle Extracto", width=250)
            },
            width="stretch",
            hide_index=True
        )

        with st.expander("🔎 Ver detalle completo / instrucción expandida por caso"):
            st.caption("Selecciona cualquier caso para ver la instrucción contable completa y clara:")
            caso_idx = st.selectbox(
                "Seleccionar movimiento:",
                options=range(len(df_display)),
                format_func=lambda i: f"Caso #{i+1} | ${df_display.iloc[i]['Valor ($)']:,.2f} | Doc: {df_display.iloc[i]['Recibo / Egreso #']} | Banco: {df_display.iloc[i]['Banco Real (Entró)']}"
            )
            sel_row = df_display.iloc[caso_idx]
            ci_1, ci_2 = st.columns(2)
            with ci_1:
                st.markdown(f"**🏦 Banco donde entró el dinero:** {sel_row['Banco Real (Entró)']}")
                st.markdown(f"**💰 Valor del movimiento:** ${sel_row['Valor ($)']:,.2f}")
                st.markdown(f"**📅 Fechas:** Banco: `{sel_row['Fecha Extracto']}` | Karing: `{sel_row['Fecha Karing']}`")
                st.markdown(f"**📄 Documento / Recibo:** `{sel_row['Recibo / Egreso #']}`")
            with ci_2:
                st.markdown(f"**❌ Cuenta errónea registrada en Karing:** {sel_row['Cuenta Errónea en Karing']}")
                st.markdown(f"**📝 Detalle en Karing:** {sel_row['Detalle en Karing']}")
                st.markdown(f"**🏦 Detalle en Extracto:** {sel_row['Detalle en Extracto']}")
            st.info(f"**📋 Instrucción exacta para el contador:**\n\n{sel_row['Instrucción de Reclasificación']}")
    else:
        st.success("🎉 ¡Excelente! No se detectaron cruces intercuentas ni registros desubicados para este mes.")

# =============================================================
# VISTA 3: LO QUE FALTA POR CONCILIAR (PENDIENTES)
# =============================================================
elif st.session_state["active_tab"] == MENU_OPTIONS[2]:
    st.markdown("""
    <div class="warning-banner">
        <h4 style="margin:0 0 6px 0; color:#78350f;">❓ ¿Qué es esta sección? (Lo que FALTA por cuadrar)</h4>
        Aquí se desglosan las partidas que no se han podido cruzar directamente. Se dividen en dos grupos claros:<br>
        1. <b>Falta en Karing (Pendientes Banco):</b> Dinero que se movió en el banco pero NO está contabilizado en Karing.<br>
        2. <b>Falta en Banco (Pendientes Karing):</b> Recibos o egresos en libros contables que NO entraron ni salieron del banco.
    </div>
    """, unsafe_allow_html=True)

    tab_f_bank, tab_f_kar = st.tabs([
        "🏦 1. Falta en Karing (Movimientos del Banco sin Contabilizar)",
        "📖 2. Falta en Banco (Registros en Karing no Reflejados en Extracto)"
    ])

    with tab_f_bank:
        st.markdown("##### 🏦 Movimientos en Extracto Bancario pendientes de registrar en Karing")
        st.write("Representan consignaciones de clientes no identificadas, transferencias entrantes sin recibo, o notas débito bancarias.")

        pb_items = []
        for acc_k, acc_r in reconciliation["cuentas"].items():
            pb_items.extend(acc_r.get("pendientes_banco", []))

        if pb_items:
            df_pb = pd.DataFrame(pb_items)
            
            # Filtro por cuenta
            ctas_pb = ["TODAS"] + sorted(list(df_pb["cuenta"].unique()))
            f_cta_pb = st.selectbox("Filtrar por Cuenta:", ctas_pb, key="f_cta_pb")
            
            df_pb_view = df_pb.copy()
            if f_cta_pb != "TODAS":
                df_pb_view = df_pb_view[df_pb_view["cuenta"] == f_cta_pb]

            # Asegurar que todas las columnas requeridas existan defensivamente
            expected_pb_cols = {
                "cuenta": "",
                "fecha": "",
                "monto": 0.0,
                "tipo": "",
                "descripcion": "",
                "documento": "",
                "contexto_monto": "Monto único en el mes"
            }
            for col, def_val in expected_pb_cols.items():
                if col not in df_pb_view.columns:
                    df_pb_view[col] = def_val

            df_pb_view["tipo"] = df_pb_view["tipo"].replace({
                "ABONO_NO_IDENTIFICADO": "Abono / Ingreso",
                "CARGO_NO_REGISTRADO": "Cargo / Egreso"
            })

            # Referencia bancaria en lenguaje humano: si el banco no genera número (común en transferencias electrónicas), indicarlo claramente
            df_pb_view["ref_banco"] = df_pb_view["documento"].fillna("").astype(str).str.strip()
            df_pb_view["ref_banco"] = df_pb_view["ref_banco"].replace({"": "(Sin ref. en extracto)", "None": "(Sin ref. en extracto)", "nan": "(Sin ref. en extracto)"})

            # Acción contable humana para que el equipo sepa qué hacer en Karing
            df_pb_view["accion_karing"] = df_pb_view["tipo"].apply(
                lambda t: "📥 Elaborar Recibo de Caja en Karing" if "Ingreso" in t or "Abono" in t else "📤 Elaborar Comprobante de Egreso en Karing"
            )

            # El extracto bancario lista todos los valores en positivo (sin signos menos);
            # el sentido del movimiento ya lo indica la columna "Tipo de Movimiento" (Abono vs Cargo)
            df_pb_view["monto"] = df_pb_view["monto"].abs()

            df_pb_disp = df_pb_view[[
                "cuenta", "fecha", "monto", "tipo", "descripcion", "ref_banco", "contexto_monto", "accion_karing"
            ]].rename(columns={
                "cuenta": "Cuenta Bancaria",
                "fecha": "Fecha en Extracto",
                "monto": "Monto ($)",
                "tipo": "Tipo de Movimiento",
                "descripcion": "Descripción en Extracto",
                "ref_banco": "No. Referencia Banco",
                "contexto_monto": "Auditoría de Montos Repetidos",
                "accion_karing": "Qué Falta Hacer en Karing"
            })

            # Botón de Descarga Excel
            excel_pb_bytes = safe_export_dataframe_to_excel_bytes(
                df_pb_disp,
                title=f"Falta en Karing (Pendientes Banco) - {selected_month}",
                sheet_name="Falta_en_Karing"
            )
            col_b1, col_b2 = st.columns([3, 1])
            with col_b1:
                st.metric("Total Partidas por Contabilizar", f"{len(df_pb_disp):,} movimientos", f"${df_pb_disp['Monto ($)'].abs().sum():,.2f}")
            with col_b2:
                st.download_button(
                    label="📥 Descargar Falta en Karing en Excel (.xlsx)",
                    data=excel_pb_bytes,
                    file_name=f"Faltantes_en_Karing_{selected_month.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch"
                )

            st.dataframe(
                df_pb_disp,
                column_config={
                    "Cuenta Bancaria": st.column_config.TextColumn("Cuenta Bancaria", width=160),
                    "Fecha en Extracto": st.column_config.TextColumn("Fecha Banco", width=105),
                    "Monto ($)": st.column_config.NumberColumn("Monto ($)", format="$ %,.2f", width=130),
                    "Tipo de Movimiento": st.column_config.TextColumn("Tipo Movimiento", width=135),
                    "Descripción en Extracto": st.column_config.TextColumn("Descripción en Extracto", width=330),
                    "No. Referencia Banco": st.column_config.TextColumn("Ref. Banco", width=150),
                    "Auditoría de Montos Repetidos": st.column_config.TextColumn("Auditoría de Repetición", width=280),
                    "Qué Falta Hacer en Karing": st.column_config.TextColumn("Qué Falta en Karing", width=250)
                },
                width="stretch",
                hide_index=True
            )
        else:
            st.success("🎉 Todas las partidas del extracto bancario están conciliadas o identificadas.")

    with tab_f_kar:
        st.markdown("##### 📖 Registros en Libros Karing no reflejados en el Extracto Bancario")
        st.write("Representan recibos de caja que nunca ingresaron a la cuenta, cheques girados no cobrados, o posibles recibos duplicados/anulados.")

        pk_items = []
        for acc_k, acc_r in reconciliation["cuentas"].items():
            pk_items.extend(acc_r.get("pendientes_karing", []))

        if pk_items:
            df_pk = pd.DataFrame(pk_items)
            
            ctas_pk = ["TODAS"] + sorted(list(df_pk["cuenta"].unique()))
            f_cta_pk = st.selectbox("Filtrar por Cuenta:", ctas_pk, key="f_cta_pk")
            
            df_pk_view = df_pk.copy()
            if f_cta_pk != "TODAS":
                df_pk_view = df_pk_view[df_pk_view["cuenta"] == f_cta_pk]

            # Asegurar que todas las columnas requeridas existan defensivamente
            expected_pk_cols = {
                "cuenta": "",
                "fecha": "",
                "debito": 0.0,
                "credito": 0.0,
                "documento": "",
                "tipo_documento": "RC",
                "tipo": "",
                "detalle": "",
                "contexto_monto": "Monto único en el mes"
            }
            for col, def_val in expected_pk_cols.items():
                if col not in df_pk_view.columns:
                    df_pk_view[col] = def_val

            df_pk_view["tipo"] = df_pk_view["tipo"].replace({
                "DEBITO_PENDIENTE_BANCO": "Ingreso en Karing (No en Extracto)",
                "CREDITO_PENDIENTE_BANCO": "Egreso en Karing (No en Extracto)"
            })

            # Acción de verificación en lenguaje humano
            df_pk_view["accion_verificacion"] = df_pk_view["tipo"].apply(
                lambda t: "Verificar si la consignación no ha ingresado al banco o anular recibo si fue devuelto" if "Ingreso" in t else "Verificar cheque girado pendiente de cobro o débito diferido"
            )

            df_pk_disp = df_pk_view[[
                "cuenta", "fecha", "debito", "credito", "documento", "tipo_documento", "tipo", "detalle", "contexto_monto", "accion_verificacion"
            ]].rename(columns={
                "cuenta": "Cuenta en Karing",
                "fecha": "Fecha Contable",
                "debito": "Débito ($)",
                "credito": "Crédito ($)",
                "documento": "No. Recibo / Doc Karing",
                "tipo_documento": "Tipo Doc (RC/CE)",
                "tipo": "Estado de la Partida",
                "detalle": "Tercero / Concepto en Karing",
                "contexto_monto": "Auditoría de Montos Repetidos",
                "accion_verificacion": "Acción Recomendada"
            })

            # Botón de Descarga Excel
            excel_pk_bytes = safe_export_dataframe_to_excel_bytes(
                df_pk_disp,
                title=f"Falta en Banco (Pendientes Karing) - {selected_month}",
                sheet_name="Falta_en_Banco"
            )
            col_k1, col_k2 = st.columns([3, 1])
            with col_k1:
                st.metric("Total Registros sin Reflejo en Banco", f"{len(df_pk_disp):,} registros", f"${(df_pk_disp['Débito ($)'] + df_pk_disp['Crédito ($)']).sum():,.2f}")
            with col_k2:
                st.download_button(
                    label="📥 Descargar Falta en Banco en Excel (.xlsx)",
                    data=excel_pk_bytes,
                    file_name=f"Faltantes_en_Banco_{selected_month.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch"
                )

            st.dataframe(
                df_pk_disp,
                column_config={
                    "Cuenta en Karing": st.column_config.TextColumn("Cuenta en Karing", width=170),
                    "Fecha Contable": st.column_config.TextColumn("Fecha Contable", width=110),
                    "Débito ($)": st.column_config.NumberColumn("Débito ($)", format="$ %,.2f", width=130),
                    "Crédito ($)": st.column_config.NumberColumn("Crédito ($)", format="$ %,.2f", width=130),
                    "No. Recibo / Doc Karing": st.column_config.TextColumn("No. Recibo / Doc", width=130),
                    "Tipo Doc (RC/CE)": st.column_config.TextColumn("Tipo Doc", width=95),
                    "Estado de la Partida": st.column_config.TextColumn("Estado", width=170),
                    "Tercero / Concepto en Karing": st.column_config.TextColumn("Detalle / Tercero", width=300),
                    "Auditoría de Montos Repetidos": st.column_config.TextColumn("Auditoría de Repetición", width=280),
                    "Acción Recomendada": st.column_config.TextColumn("Acción Recomendada", width=300)
                },
                width="stretch",
                hide_index=True
            )
        else:
            st.success("🎉 Todos los registros del libro auxiliar de Karing se reflejaron en el extracto bancario.")

# =============================================================
# VISTA 4: GASTOS E IMPUESTOS BANCARIOS (4X1000)
# =============================================================
elif st.session_state["active_tab"] == MENU_OPTIONS[3]:
    st.markdown("""
    <div class="info-banner">
        <h4 style="margin:0 0 6px 0; color:#1e293b;">💸 Gastos Bancarios, GMF (4x1000), Comisiones e Intereses</h4>
        Movimientos automáticos cobrados o abonados por los bancos. 
        Permite filtrar y descargar en Excel <b>por Cuenta Bancaria</b> y <b>por Periodo / Mes</b>, 
        con su respectivo asiento contable modelo sugerido para asentar en Karing.
    </div>
    """, unsafe_allow_html=True)

    # 1. Controles de Filtrado: Mes / Periodo, Cuenta Bancaria y Concepto
    c_f1, c_f2, c_f3 = st.columns([1.5, 2.0, 1.3])
    with c_f1:
        mes_opts = [f"Mes Seleccionado ({selected_month})"] + [m for m in available_months if m != selected_month] + ["📊 TODOS LOS MESES (Consolidado Histórico)"]
        sel_mes_opt = st.selectbox("📅 Periodo / Mes a Consultar y Descargar:", options=mes_opts, key="flt_exp_mes")

    # Determinar qué meses cargar
    if "TODOS LOS MESES" in sel_mes_opt:
        target_months = available_months
        periodo_title = "Todos los Meses (Consolidado)"
        periodo_slug = "TODOS_LOS_MESES"
    elif "Mes Seleccionado" in sel_mes_opt:
        target_months = [selected_month]
        periodo_title = selected_month
        periodo_slug = selected_month.replace(" ", "_")
    else:
        target_months = [sel_mes_opt]
        periodo_title = sel_mes_opt
        periodo_slug = sel_mes_opt.replace(" ", "_")

    # 2. Cargar gastos e impuestos del periodo solicitado
    all_exp = []
    for tm in target_months:
        rec_tm = run_reconciliation(tm, tolerance_days)
        for acc_k, acc_r in rec_tm["cuentas"].items():
            for g in acc_r.get("gastos_impuestos", []):
                item_g = g.copy()
                item_g["periodo"] = tm
                all_exp.append(item_g)

    if all_exp:
        # Construir nombres amigables de cuentas
        all_account_names = sorted(list(set(g["cuenta_banco"] for g in all_exp if g.get("cuenta_banco"))))
        account_labels = {}
        for an in all_account_names:
            matched_cfg = None
            for cfg in ACCOUNTS_CATALOG.values():
                if cfg.karing_name == an:
                    matched_cfg = cfg
                    break
            if matched_cfg:
                account_labels[an] = f"{matched_cfg.bank_name} - {matched_cfg.karing_name} (#{str(matched_cfg.account_number)[-4:]})"
            else:
                account_labels[an] = an

        ctas_options = ["TODAS LAS CUENTAS"] + all_account_names

        with c_f2:
            sel_cta = st.selectbox(
                "🏦 Filtrar por Cuenta Bancaria:",
                options=ctas_options,
                format_func=get_account_label,
                key="flt_exp_cta"
            )

        with c_f3:
            tipos_disp = ["TODOS LOS CONCEPTOS", "GMF (4x1000)", "Comisiones Bancarias", "Rendimientos / Intereses", "IVA / Otros"]
            sel_tipo = st.selectbox("📑 Filtrar por Concepto:", options=tipos_disp, key="flt_exp_tipo")

        # 3. Aplicar filtros
        filtered_exp = all_exp.copy()
        if sel_cta != "TODAS LAS CUENTAS":
            filtered_exp = [g for g in filtered_exp if g.get("cuenta_banco") == sel_cta]

        if sel_tipo == "GMF (4x1000)":
            filtered_exp = [g for g in filtered_exp if "GMF" in g.get("tipo", "").upper()]
        elif sel_tipo == "Comisiones Bancarias":
            filtered_exp = [g for g in filtered_exp if "COMISION" in g.get("tipo", "").upper()]
        elif sel_tipo == "Rendimientos / Intereses":
            filtered_exp = [g for g in filtered_exp if "INTERES" in g.get("tipo", "").upper() or "RENDIMIENTO" in g.get("tipo", "").upper()]
        elif sel_tipo == "IVA / Otros":
            filtered_exp = [g for g in filtered_exp if g.get("tipo", "").upper() not in ["GMF_4X1000", "COMISION", "COMISION_BOLD", "INTERESES", "RENDIMIENTO"]]

        # 4. Métricas de Conceptos Dinámicas
        calc_gmf = sum(abs(g["monto"]) for g in filtered_exp if "GMF" in g.get("tipo", "").upper())
        calc_com = sum(abs(g["monto"]) for g in filtered_exp if "COMISION" in g.get("tipo", "").upper())
        calc_int = sum(abs(g["monto"]) for g in filtered_exp if "INTERES" in g.get("tipo", "").upper() or "RENDIMIENTO" in g.get("tipo", "").upper())
        calc_tot_gastos = sum(abs(g["monto"]) for g in filtered_exp if g.get("monto", 0) < 0)

        m_g1, m_g2, m_g3, m_g4 = st.columns(4)
        with m_g1:
            st.metric("GMF (4x1000)", f"${calc_gmf:,.2f}")
        with m_g2:
            st.metric("Comisiones ACH / Portales / Bold", f"${calc_com:,.2f}")
        with m_g3:
            st.metric("Intereses a Favor (Rendimientos)", f"${calc_int:,.2f}")
        with m_g4:
            st.metric("Total Gastos e Impuestos", f"${calc_tot_gastos:,.2f}")

        # 5. Plantilla de Asiento Contable Dinámica
        st.markdown("#### 📝 Asiento Contable Modelo Sugerido para Contabilizar en Karing:")
        
        if sel_cta != "TODAS LAS CUENTAS":
            matched_cfg = None
            for cfg in ACCOUNTS_CATALOG.values():
                if cfg.karing_name == sel_cta:
                    matched_cfg = cfg
                    break
            cta_code_str = f"{matched_cfg.karing_code} - {matched_cfg.karing_name}" if matched_cfg else f"11XXXX - {sel_cta}"
            cuenta_label_asiento = f"Cuenta: {sel_cta}"
        else:
            cta_code_str = "11XXXX - Bancos (Varias Cuentas)"
            cuenta_label_asiento = "Todas las Cuentas Bancarias"

        st.code(f"""
-- ASIENTO DE AJUSTE BANCARIO - {periodo_title} ({cuenta_label_asiento})
---------------------------------------------------------------------------------------------
CUENTA CONTABLE                    CONCEPTO                      DÉBITO ($)       CRÉDITO ($)
---------------------------------------------------------------------------------------------
511595 - GMF 4x1000                Impuesto GMF del periodo      ${calc_gmf:,.2f}
530515 - Comisiones Bancarias      Comisiones ACH / Portal/Bold  ${calc_com:,.2f}
{cta_code_str:<34} Salida por gastos bancarios                     ${calc_tot_gastos:,.2f}
---------------------------------------------------------------------------------------------
{cta_code_str:<34} Entrada por rendimientos      ${calc_int:,.2f}
421005 - Ingresos Financieros      Intereses abonados periodo                     ${calc_int:,.2f}
---------------------------------------------------------------------------------------------
        """, language="sql")

        # 6. Preparar DataFrame para visualización y descarga
        st.markdown("#### Detalle Individual de Cada Cargo:")
        df_exp_disp = pd.DataFrame(filtered_exp)
        
        # Asegurar columnas defensivamente
        for col, def_val in [
            ("periodo", periodo_title),
            ("banco", ""),
            ("cuenta_banco", ""),
            ("fecha", ""),
            ("monto", 0.0),
            ("tipo", ""),
            ("documento", ""),
            ("sugerencia_contable", ""),
            ("descripcion", "")
        ]:
            if col not in df_exp_disp.columns:
                df_exp_disp[col] = def_val

        df_exp_disp["documento"] = df_exp_disp["documento"].fillna("").astype(str).str.strip().replace({"": "(Sin ref. en extracto)", "None": "(Sin ref. en extracto)", "nan": "(Sin ref. en extracto)"})
        
        df_exp_disp["tipo_humano"] = df_exp_disp["tipo"].replace({
            "GMF_4X1000": "GMF (4x1000)",
            "COMISION": "Comisión Bancaria / ACH",
            "COMISION_BOLD": "Comisión Pasarela Bold",
            "INTERESES": "Intereses a Favor",
            "IVA": "IVA Bancario"
        })

        cols_to_use = [
            "periodo", "banco", "cuenta_banco", "fecha", "monto", "tipo_humano", "documento", "sugerencia_contable", "descripcion"
        ]
        
        df_exp_final = df_exp_disp[cols_to_use].rename(columns={
            "periodo": "Periodo",
            "banco": "Banco",
            "cuenta_banco": "Cuenta Bancaria",
            "fecha": "Fecha",
            "monto": "Monto ($)",
            "tipo_humano": "Tipo de Partida",
            "documento": "No. Referencia Banco",
            "sugerencia_contable": "Asiento Contable Sugerido",
            "descripcion": "Descripción en Extracto"
        })

        # 7. Botón de Descarga Excel personalizado
        cta_slug = "Todas_Cuentas" if sel_cta == "TODAS LAS CUENTAS" else re.sub(r'[^a-zA-Z0-9_]+', '_', sel_cta).strip('_')
        file_name_excel = f"Gastos_Impuestos_{cta_slug}_{periodo_slug}.xlsx"
        sheet_name_excel = f"Gastos_{cta_slug[:18]}"

        excel_exp_bytes = safe_export_dataframe_to_excel_bytes(
            df_exp_final,
            title=f"Gastos e Impuestos Bancarios - {sel_cta} ({periodo_title})",
            sheet_name=sheet_name_excel
        )

        col_ex1, col_ex2 = st.columns([2.0, 2.0])
        with col_ex1:
            st.caption(f"Mostrando **{len(df_exp_final):,}** movimientos para **{get_account_label(sel_cta)}** en **{periodo_title}**.")
        with col_ex2:
            cta_btn_desc = "Todas las Cuentas" if sel_cta == "TODAS LAS CUENTAS" else (sel_cta[:22] + "...")
            st.download_button(
                label=f"📥 Descargar en Excel ({cta_btn_desc})",
                data=excel_exp_bytes,
                file_name=file_name_excel,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch"
            )

        # 8. Tabla Interactiva
        st.dataframe(
            df_exp_final,
            column_config={
                "Periodo": st.column_config.TextColumn("Periodo", width=110),
                "Banco": st.column_config.TextColumn("Banco", width=115),
                "Cuenta Bancaria": st.column_config.TextColumn("Cuenta Bancaria", width=165),
                "Fecha": st.column_config.TextColumn("Fecha", width=100),
                "Monto ($)": st.column_config.NumberColumn("Monto ($)", format="$ %,.2f", width=125),
                "Tipo de Partida": st.column_config.TextColumn("Concepto", width=130),
                "No. Referencia Banco": st.column_config.TextColumn("Ref. Banco", width=120),
                "Asiento Contable Sugerido": st.column_config.TextColumn("Asiento Contable Sugerido", width=290),
                "Descripción en Extracto": st.column_config.TextColumn("Descripción en Extracto", width=320)
            },
            width="stretch",
            hide_index=True
        )
    else:
        st.success("🎉 No se registraron gastos bancarios ni impuestos en la selección consultada.")

# =============================================================
# VISTA 5: INTERESES Y RENDIMIENTOS DE CUENTAS (INGRESOS FINANCIEROS)
# =============================================================
elif st.session_state["active_tab"] == MENU_OPTIONS[4]:
    st.markdown("""
    <div class="info-banner" style="border-left: 5px solid #0284c7;">
        <h4 style="margin:0 0 6px 0; color:#0369a1;">💰 ¿Qué es esta sección? (Intereses y Rendimientos Financieros Abonados)</h4>
        <b>Naturaleza Contable:</b> Son dineros que las entidades financieras (Bancolombia, Banco de Bogotá, Davivienda, etc.) abonaron directamente en las cuentas bancarias de Planetour por concepto de <b>rendimientos diarios de cuentas de ahorro, inversiones virtuales o liquidación de CDTs</b>.<br>
        <b>Estado de Conciliación:</b> Los extractos bancarios los reflejan plenamente, pero en <b>Karing actualmente no están causados (0 registros en libros contables)</b>.<br>
        <b>Acción Contable Requerida:</b> Generar la Nota de Contabilidad debitando la cuenta bancaria donde ingresó el dinero y acreditando la cuenta de <b>Ingresos Financieros (421005)</b> para que los saldos contables suban y cuadren exactamente con los extractos.
    </div>
    """, unsafe_allow_html=True)

    def get_month_interests(m_name: str):
        m_rec = run_reconciliation(m_name, tolerance_days)
        m_ints = []
        for acc_k, acc_r in m_rec.get("cuentas", {}).items():
            for g in acc_r.get("gastos_impuestos", []):
                if g.get("tipo") == "INTERESES" or "RENDIMIENTO" in str(g.get("descripcion", "")).upper():
                    item = dict(g)
                    item["periodo"] = m_name
                    m_ints.append(item)
        return m_ints

    period_options_int = [f"MES ACTUAL ({selected_month})"]
    for m in available_months:
        if m != selected_month:
            period_options_int.append(f"Periodo: {m}")
    if len(available_months) > 1:
        period_options_int.append("TODOS LOS MESES (Consolidado Histórico)")

    st.markdown("##### 🔍 Filtros de Consulta y Descarga:")
    f_int1, f_int2, f_int3 = st.columns(3)

    with f_int1:
        sel_per_choice = st.selectbox("📅 Seleccionar Periodo / Mes:", options=period_options_int, key="flt_int_per")

    if sel_per_choice.startswith("MES ACTUAL"):
        all_interests = get_month_interests(selected_month)
        periodo_int_title = selected_month
    elif sel_per_choice == "TODOS LOS MESES (Consolidado Histórico)":
        all_interests = []
        for m in available_months:
            all_interests.extend(get_month_interests(m))
        periodo_int_title = "CONSOLIDADO_HISTORICO"
    else:
        chosen_m = sel_per_choice.replace("Periodo: ", "").strip()
        all_interests = get_month_interests(chosen_m)
        periodo_int_title = chosen_m

    with f_int2:
        ctas_with_interest = sorted(list({i.get("cuenta_banco") for i in all_interests if i.get("cuenta_banco")}))
        ctas_int_options = ["TODAS LAS CUENTAS"] + ctas_with_interest
        sel_cta_int = st.selectbox(
            "🏦 Filtrar por Cuenta Bancaria:",
            options=ctas_int_options,
            format_func=get_account_label,
            key="flt_int_cta"
        )

    with f_int3:
        tipos_int_disp = ["TODOS LOS RENDIMIENTOS", "Abono Intereses Cuentas Ahorro", "Inversión Virtual / CDTs"]
        sel_tipo_int = st.selectbox("📑 Filtrar por Tipo de Rendimiento:", options=tipos_int_disp, key="flt_int_tipo")

    filtered_int = all_interests.copy()
    if sel_cta_int != "TODAS LAS CUENTAS":
        filtered_int = [i for i in filtered_int if i.get("cuenta_banco") == sel_cta_int]

    if sel_tipo_int == "Abono Intereses Cuentas Ahorro":
        filtered_int = [i for i in filtered_int if "AHORRO" in str(i.get("descripcion", "")).upper() or "RENDIMIENTOS FINANCIEROS" in str(i.get("descripcion", "")).upper()]
    elif sel_tipo_int == "Inversión Virtual / CDTs":
        filtered_int = [i for i in filtered_int if any(k in str(i.get("descripcion", "")).upper() for k in ["CDT", "VIRTUAL", "INV"])]

    total_val_int = sum(abs(float(i.get("monto", 0.0))) for i in filtered_int)
    num_abonos_int = len(filtered_int)
    max_abono_int = max([abs(float(i.get("monto", 0.0))) for i in filtered_int], default=0.0)

    m_i1, m_i2, m_i3, m_i4 = st.columns(4)
    with m_i1:
        st.metric("Total Rendimientos Abonados ($)", f"${total_val_int:,.2f}")
    with m_i2:
        st.metric("Número de Abonos Registrados", f"{num_abonos_int:,} abonos")
    with m_i3:
        st.metric("Mayor Rendimiento Unitario", f"${max_abono_int:,.2f}")
    with m_i4:
        st.metric("Estado en Libros Karing", "🔴 0 causados (Pendiente)")

    st.markdown("#### 📝 Asiento Contable Sugerido para Causar en Karing (Ingresos Financieros):")

    if sel_cta_int != "TODAS LAS CUENTAS":
        matched_cfg_int = None
        for cfg in ACCOUNTS_CATALOG.values():
            if cfg.karing_name == sel_cta_int:
                matched_cfg_int = cfg
                break
        cta_code_int = f"{matched_cfg_int.karing_code} - {matched_cfg_int.karing_name}" if matched_cfg_int else f"11XXXX - {sel_cta_int}"
        cuenta_label_int = f"Cuenta: {sel_cta_int}"
    else:
        cta_code_int = "11XXXX - Bancos (Varias Cuentas)"
        cuenta_label_int = "Todas las Cuentas con Rendimientos"

    st.code(f"""
-- ASIENTO DE CAUSACIÓN DE RENDIMIENTOS FINANCIEROS - {periodo_int_title} ({cuenta_label_int})
---------------------------------------------------------------------------------------------------
CÓDIGO CONTABLE                      DESCRIPCIÓN                        DÉBITO ($)       CRÉDITO ($)
---------------------------------------------------------------------------------------------------
{cta_code_int:<36} Entrada de rendimientos al banco   ${total_val_int:,.2f}
421005 - Ingresos Financieros        Rendimientos e intereses periodo                    ${total_val_int:,.2f}
---------------------------------------------------------------------------------------------------
SUMAS IGUALES                                                      ${total_val_int:,.2f}     ${total_val_int:,.2f}
    """, language="sql")

    if sel_cta_int == "TODAS LAS CUENTAS" and filtered_int:
        with st.expander("📊 Ver desglose de causación por cada cuenta bancaria individual:"):
            st.caption("Distribución exacta para elaborar el comprobante contable en Karing:")
            breakdown_dict = {}
            for item in filtered_int:
                cta = item.get("cuenta_banco", "")
                val = abs(float(item.get("monto", 0.0)))
                breakdown_dict[cta] = breakdown_dict.get(cta, 0.0) + val

            bd_rows = []
            for cta, val in sorted(breakdown_dict.items(), key=lambda x: x[1], reverse=True):
                matched_cfg = next((c for c in ACCOUNTS_CATALOG.values() if c.karing_name == cta), None)
                code_str = f"{matched_cfg.karing_code} - {matched_cfg.karing_name}" if matched_cfg else cta
                bd_rows.append({
                    "Cuenta Contable Karing": code_str,
                    "Total Rendimiento ($)": val,
                    "Participación (%)": f"{(val / total_val_int * 100):.1f}%" if total_val_int > 0 else "0%"
                })
            st.dataframe(pd.DataFrame(bd_rows), width="stretch", hide_index=True)

    st.markdown("#### Detalle Individual de Cada Abono de Interés en Extracto:")

    if filtered_int:
        df_int_disp = pd.DataFrame(filtered_int)
        for col, def_val in [
            ("periodo", periodo_int_title),
            ("banco", ""),
            ("cuenta_banco", ""),
            ("fecha", ""),
            ("monto", 0.0),
            ("descripcion", ""),
            ("sugerencia_contable", "")
        ]:
            if col not in df_int_disp.columns:
                df_int_disp[col] = def_val

        df_int_disp["monto"] = df_int_disp["monto"].abs()

        df_int_final = df_int_disp[[
            "periodo", "banco", "cuenta_banco", "fecha", "monto", "descripcion", "sugerencia_contable"
        ]].rename(columns={
            "periodo": "Periodo",
            "banco": "Banco",
            "cuenta_banco": "Cuenta Bancaria",
            "fecha": "Fecha de Abono",
            "monto": "Valor Rendimiento ($)",
            "descripcion": "Descripción en Extracto",
            "sugerencia_contable": "Asiento Contable Sugerido"
        })

        excel_int_bytes = safe_export_dataframe_to_excel_bytes(
            df_int_final,
            title=f"Rendimientos e Intereses Bancarios - {periodo_int_title}",
            sheet_name="Intereses_Rendimientos"
        )
        col_di1, col_di2 = st.columns([3, 1])
        with col_di2:
            fname_clean_per = periodo_int_title.replace(' ', '_').replace('(', '').replace(')', '')
            fname_clean_cta = sel_cta_int.replace(' ', '_')[:20]
            st.download_button(
                label="📥 Descargar Listado en Excel (.xlsx)",
                data=excel_int_bytes,
                file_name=f"Intereses_Rendimientos_{fname_clean_per}_{fname_clean_cta}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch"
            )

        st.dataframe(
            df_int_final,
            column_config={
                "Periodo": st.column_config.TextColumn("Periodo", width=110),
                "Banco": st.column_config.TextColumn("Banco", width=115),
                "Cuenta Bancaria": st.column_config.TextColumn("Cuenta Bancaria", width=170),
                "Fecha de Abono": st.column_config.TextColumn("Fecha Abono", width=105),
                "Valor Rendimiento ($)": st.column_config.NumberColumn("Valor Rendimiento ($)", format="$ %,.2f", width=150),
                "Descripción en Extracto": st.column_config.TextColumn("Descripción en Extracto", width=340),
                "Asiento Contable Sugerido": st.column_config.TextColumn("Asiento Contable Sugerido", width=380)
            },
            width="stretch",
            hide_index=True
        )
    else:
        st.success("🎉 No se registraron abonos de intereses en la selección consultada.")

# =============================================================
# VISTA 6: LO QUE ESTÁ BIEN (CONCILIADOS 1 A 1)
# =============================================================
elif st.session_state["active_tab"] == MENU_OPTIONS[5]:
    st.markdown("""
    <div class="info-banner">
        <h4 style="margin:0 0 6px 0; color:#1e293b;">✅ Movimientos Conciliados Exitosamente (1 a 1)</h4>
        Partidas que coinciden perfectamente en valor, sentido (ingreso o egreso) y proximidad de fechas entre el extracto bancario y Karing.
    </div>
    """, unsafe_allow_html=True)

    all_mat = []
    for acc_k, acc_r in reconciliation["cuentas"].items():
        all_mat.extend(acc_r.get("conciliados", []))

    if all_mat:
        df_mat = pd.DataFrame(all_mat)

        c_m1, c_m2 = st.columns(2)
        with c_m1:
            st.metric("Total Partidas Conciliadas", f"{len(df_mat):,} movimientos")
        with c_m2:
            st.metric("Monto Total Conciliado", f"${df_mat['valor_conciliado'].sum():,.2f}")

        # Asegurar columnas de pareo explícito
        if "id_pareo" not in df_mat.columns:
            df_mat["id_pareo"] = [f"PAR-{i+1:04d}" for i in range(len(df_mat))]
        if "documento_banco" not in df_mat.columns:
            df_mat["documento_banco"] = ""
        if "tipo_documento_karing" not in df_mat.columns:
            df_mat["tipo_documento_karing"] = "RC"

        df_mat["documento_banco"] = df_mat["documento_banco"].fillna("").astype(str).str.strip().replace({"": "(Sin ref)", "None": "(Sin ref)", "nan": "(Sin ref)"})

        # Filtros rápidos en Conciliados
        st.markdown("##### 🔍 Filtros de Búsqueda de Pareos:")
        fc1, fc2, fc3 = st.columns([1.5, 1.5, 1])
        with fc1:
            ctas_mat = ["TODAS"] + sorted(list(df_mat["cuenta_karing"].unique()))
            sel_cta_mat = st.selectbox("Filtrar por Cuenta:", ctas_mat, key="sel_cta_mat")
        with fc2:
            search_doc_mat = st.text_input("Buscar por Recibo / Doc Karing o Ref Banco:", key="search_doc_mat")
        with fc3:
            filter_amt_mat = st.number_input("Filtrar por Monto Exacto ($):", min_value=0.0, step=1000.0, value=0.0, key="amt_mat")

        df_mat_filtered = df_mat.copy()
        if sel_cta_mat != "TODAS":
            df_mat_filtered = df_mat_filtered[df_mat_filtered["cuenta_karing"] == sel_cta_mat]
        if search_doc_mat.strip():
            sd = search_doc_mat.strip().lower()
            df_mat_filtered = df_mat_filtered[
                df_mat_filtered["documento_karing"].astype(str).str.lower().str.contains(sd) |
                df_mat_filtered["documento_banco"].astype(str).str.lower().str.contains(sd) |
                df_mat_filtered["detalle_karing"].astype(str).str.lower().str.contains(sd) |
                df_mat_filtered["descripcion_banco"].astype(str).str.lower().str.contains(sd)
            ]
        if filter_amt_mat > 0:
            df_mat_filtered = df_mat_filtered[df_mat_filtered["valor_conciliado"].round(2) == round(filter_amt_mat, 2)]

        df_mat_disp = df_mat_filtered[[
            "id_pareo", "cuenta_karing", "valor_conciliado", "sentido", "fecha_banco", "documento_banco",
            "descripcion_banco", "fecha_karing", "documento_karing", "tipo_documento_karing", "detalle_karing", "diferencia_dias"
        ]].rename(columns={
            "id_pareo": "ID Pareo",
            "cuenta_karing": "Cuenta",
            "valor_conciliado": "Valor Conciliado ($)",
            "sentido": "Sentido",
            "fecha_banco": "Fecha Banco",
            "documento_banco": "Ref. Banco",
            "descripcion_banco": "Descripción en Extracto",
            "fecha_karing": "Fecha Karing",
            "documento_karing": "Doc Karing",
            "tipo_documento_karing": "Tipo Doc",
            "detalle_karing": "Detalle en Karing",
            "diferencia_dias": "Dif. Días"
        })

        # Botón de Descarga Excel
        excel_mat_bytes = safe_export_dataframe_to_excel_bytes(
            df_mat_disp,
            title=f"Partidas Conciliadas 1 a 1 - {selected_month}",
            sheet_name="Conciliados"
        )
        col_m1, col_m2 = st.columns([3, 1])
        with col_m1:
            st.caption(f"Mostrando **{len(df_mat_disp):,}** pareos conciliados según los filtros.")
        with col_m2:
            st.download_button(
                label="📥 Descargar Partidas Conciliadas en Excel (.xlsx)",
                data=excel_mat_bytes,
                file_name=f"Partidas_Conciliadas_{selected_month.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch"
            )

        st.dataframe(
            df_mat_disp,
            column_config={
                "ID Pareo": st.column_config.TextColumn("ID Pareo", width=120),
                "Cuenta": st.column_config.TextColumn("Cuenta", width=150),
                "Valor Conciliado ($)": st.column_config.NumberColumn("Valor ($)", format="$ %,.2f", width=130),
                "Sentido": st.column_config.TextColumn("Sentido", width=95),
                "Fecha Banco": st.column_config.TextColumn("Fecha Banco", width=105),
                "Ref. Banco": st.column_config.TextColumn("Ref. Banco", width=110),
                "Descripción en Extracto": st.column_config.TextColumn("Descripción Extracto", width=250),
                "Fecha Karing": st.column_config.TextColumn("Fecha Karing", width=105),
                "Doc Karing": st.column_config.TextColumn("Doc Karing", width=110),
                "Tipo Doc": st.column_config.TextColumn("Tipo", width=75),
                "Detalle en Karing": st.column_config.TextColumn("Detalle Karing", width=250),
                "Dif. Días": st.column_config.NumberColumn("Dif. Días", format="%d d", width=85)
            },
            width="stretch",
            hide_index=True
        )

# =============================================================
# VISTA 7: BUSCADOR UNIVERSAL
# =============================================================
elif st.session_state["active_tab"] == MENU_OPTIONS[6]:
    tab_search_gen, tab_auditoria_montos = st.tabs([
        "🔍 1. Buscador Universal (por Documento, Tercero, Texto o Monto)",
        "⚖️ 2. Auditoría y Trazabilidad de Montos Repetidos (¿Cuál macheó y cuál faltó?)"
    ])

    with tab_search_gen:
        st.markdown("""
        <div class="info-banner">
            <h4 style="margin:0 0 6px 0; color:#1e293b;">🔍 Buscador Universal de Movimientos</h4>
            Busca cualquier partida por <b>número de recibo/egreso</b>, <b>monto</b>, <b>tercero</b> o <b>concepto</b> en todas las cuentas del mes.
        </div>
        """, unsafe_allow_html=True)

        # Construir índice unificado
        unified_items = []
        for acc_k, acc_r in reconciliation["cuentas"].items():
            kn = acc_r["config"].karing_name

            for c in acc_r.get("conciliados", []):
                unified_items.append({
                    "Estado": "🟢 Conciliado 1 a 1",
                    "Fecha": c["fecha_banco"],
                    "Doc Karing": str(c["documento_karing"]),
                    "Cuenta": kn,
                    "Detalle": f"Karing: {c['detalle_karing']} | Banco: {c['descripcion_banco']}",
                    "Monto ($)": c["valor_conciliado"],
                    "Sentido": c["sentido"],
                    "Diagnóstico": f"Conciliado (ID: {c.get('id_pareo', 'PAR')}, dif: {c['diferencia_dias']} d)"
                })

            for cr in acc_r.get("cruces_intercuentas", []):
                unified_items.append({
                    "Estado": "🟡 Cruce Intercuentas (Cuenta Errónea)",
                    "Fecha": cr["fecha_banco"],
                    "Doc Karing": str(cr["documento_karing"]),
                    "Cuenta": f"Entró a {cr['banco_real']} pero registrado en {cr['cuenta_registrada_karing']}",
                    "Detalle": f"Extracto: {cr['descripcion_banco']} | Karing: {cr['detalle_karing']}",
                    "Monto ($)": cr["valor"],
                    "Sentido": cr["sentido"],
                    "Diagnóstico": cr["accion_recomendada"]
                })

            for g in acc_r.get("gastos_impuestos", []):
                unified_items.append({
                    "Estado": f"🟣 Gasto/Impuesto ({g['tipo']})",
                    "Fecha": g["fecha"],
                    "Doc Karing": str(g.get("documento", "")),
                    "Cuenta": kn,
                    "Detalle": g["descripcion"],
                    "Monto ($)": abs(g["monto"]),
                    "Sentido": "EGRESO" if g["monto"] < 0 else "INGRESO",
                    "Diagnóstico": g.get("sugerencia_contable", "Gasto bancario por contabilizar")
                })

            for pb in acc_r.get("pendientes_banco", []):
                unified_items.append({
                    "Estado": "🔴 Falta en Karing (Pendiente Banco)",
                    "Fecha": pb["fecha"],
                    "Doc Karing": str(pb.get("documento", "")),
                    "Cuenta": kn,
                    "Detalle": pb["descripcion"],
                    "Monto ($)": abs(pb["monto"]),
                    "Sentido": "INGRESO" if pb.get("tipo") == "ABONO_NO_IDENTIFICADO" else "EGRESO",
                    "Diagnóstico": f"Movimiento en extracto sin Karing. {pb.get('contexto_monto', '')}"
                })

            for pk in acc_r.get("pendientes_karing", []):
                is_deb = pk["debito"] > 0
                val = pk["debito"] if is_deb else pk["credito"]
                unified_items.append({
                    "Estado": "🟠 Falta en Banco (Pendiente Karing)",
                    "Fecha": pk["fecha"],
                    "Doc Karing": str(pk["documento"]),
                    "Cuenta": kn,
                    "Detalle": pk["detalle"],
                    "Monto ($)": val,
                    "Sentido": "INGRESO" if is_deb else "EGRESO",
                    "Diagnóstico": f"Registrado en Karing sin reflejo en banco. {pk.get('contexto_monto', '')}"
                })

        df_unified = pd.DataFrame(unified_items)

        col_q1, col_q2, col_q3 = st.columns([2, 1, 1])
        with col_q1:
            query_text = st.text_input(
                "🔎 Escribe número de recibo, tercero o palabra clave:",
                placeholder="Ej: 112511, 31099, WALDOR, NEQUI, BRE-B, TRANSFERENCIA..."
            )
        with col_q2:
            filtro_estado = st.selectbox(
                "Filtrar por Estado:",
                options=["TODOS"] + sorted(list(df_unified["Estado"].unique()))
            )
        with col_q3:
            filtro_cuenta = st.selectbox(
                "Filtrar por Cuenta:",
                options=["TODAS"] + sorted([cfg.karing_name for cfg in ACCOUNTS_CATALOG.values()])
            )

        df_filtered = df_unified.copy()
        if query_text.strip():
            q = query_text.strip().lower()
            df_filtered = df_filtered[
                df_filtered["Doc Karing"].str.lower().str.contains(q, na=False) |
                df_filtered["Detalle"].str.lower().str.contains(q, na=False) |
                df_filtered["Cuenta"].str.lower().str.contains(q, na=False) |
                df_filtered["Monto ($)"].astype(str).str.contains(q, na=False)
            ]
        if filtro_estado != "TODOS":
            df_filtered = df_filtered[df_filtered["Estado"] == filtro_estado]
        if filtro_cuenta != "TODAS":
            df_filtered = df_filtered[df_filtered["Cuenta"].str.contains(filtro_cuenta, na=False)]

        cols_search = ["Estado", "Cuenta", "Fecha", "Monto ($)", "Sentido", "Doc Karing", "Detalle", "Diagnóstico"]
        df_filtered = df_filtered[[c for c in cols_search if c in df_filtered.columns]]

        st.markdown(f"**Resultados encontrados: {len(df_filtered):,} registros**")

        # Botón Descarga Excel de la búsqueda
        excel_search_bytes = safe_export_dataframe_to_excel_bytes(
            df_filtered,
            title=f"Resultados de Busqueda - {selected_month}",
            sheet_name="Busqueda"
        )
        col_sb1, col_sb2 = st.columns([3, 1])
        with col_sb2:
            st.download_button(
                label="📥 Descargar Búsqueda en Excel (.xlsx)",
                data=excel_search_bytes,
                file_name=f"Busqueda_{selected_month.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch"
            )

        st.dataframe(
            df_filtered,
            column_config={
                "Estado": st.column_config.TextColumn("Estado", width=180),
                "Cuenta": st.column_config.TextColumn("Cuenta", width=160),
                "Fecha": st.column_config.TextColumn("Fecha", width=105),
                "Monto ($)": st.column_config.NumberColumn("Monto ($)", format="$ %,.2f", width=130),
                "Sentido": st.column_config.TextColumn("Sentido", width=95),
                "Doc Karing": st.column_config.TextColumn("No. Recibo / Ref", width=110),
                "Detalle": st.column_config.TextColumn("Detalle", width=320),
                "Diagnóstico": st.column_config.TextColumn("Diagnóstico", width=340)
            },
            width="stretch",
            hide_index=True
        )

    with tab_auditoria_montos:
        st.markdown("""
        <div class="info-banner">
            <h4 style="margin:0 0 6px 0; color:#1e293b;">⚖️ Auditoría de Montos Repetidos y Trazabilidad de Cruces</h4>
            Cuando en el extracto o en Karing existen múltiples movimientos con el mismo valor (ej. varios pagos de $3.000.000), 
            esta herramienta te permite saber <b>exactamente cuál movimiento del banco se emparejó con cuál recibo de Karing</b>, 
            y <b>cuál partida específica fue la que quedó huérfana sin machear</b>.
        </div>
        """, unsafe_allow_html=True)

        # 1. Recopilar datos globales de todas las cuentas
        all_conc_list = []
        all_pb_list = []
        all_pk_list = []

        for acc_k, acc_r in reconciliation["cuentas"].items():
            kn = acc_r["config"].karing_name
            for c in acc_r.get("conciliados", []):
                item_c = c.copy()
                item_c["cuenta_nombre"] = kn
                all_conc_list.append(item_c)
            for pb in acc_r.get("pendientes_banco", []):
                item_pb = pb.copy()
                item_pb["cuenta_nombre"] = kn
                all_pb_list.append(item_pb)
            for pk in acc_r.get("pendientes_karing", []):
                item_pk = pk.copy()
                item_pk["cuenta_nombre"] = kn
                all_pk_list.append(item_pk)

        # 2. Análisis de frecuencias de montos
        from collections import defaultdict
        amt_stats = defaultdict(lambda: {"banco_tot": 0, "karing_tot": 0, "conc": 0, "pb": 0, "pk": 0})

        for c in all_conc_list:
            v = round(c["valor_conciliado"], 2)
            amt_stats[v]["conc"] += 1
            amt_stats[v]["banco_tot"] += 1
            amt_stats[v]["karing_tot"] += 1

        for pb in all_pb_list:
            v = round(abs(pb["monto"]), 2)
            amt_stats[v]["pb"] += 1
            amt_stats[v]["banco_tot"] += 1

        for pk in all_pk_list:
            v = round(pk["debito"] if pk["debito"] > 0 else pk["credito"], 2)
            amt_stats[v]["pk"] += 1
            amt_stats[v]["karing_tot"] += 1

        # Ordenar montos: primero los que tienen discrepancias (pb > 0 o pk > 0) y repeticiones
        sorted_amts = sorted(
            amt_stats.keys(),
            key=lambda x: (
                1 if (amt_stats[x]["pb"] > 0 or amt_stats[x]["pk"] > 0) and (amt_stats[x]["banco_tot"] > 1 or amt_stats[x]["karing_tot"] > 1) else 0,
                amt_stats[x]["banco_tot"] + amt_stats[x]["karing_tot"]
            ),
            reverse=True
        )

        # Opciones para el selectbox
        amt_options = []
        amt_map = {}
        for a in sorted_amts:
            s = amt_stats[a]
            if s["banco_tot"] > 1 or s["karing_tot"] > 1:
                disp_str = f"${a:,.2f}  |  En Banco: {s['banco_tot']} vs Karing: {s['karing_tot']}  ({s['conc']} cruzaron | {s['pb']} pend. banco, {s['pk']} pend. karing)"
            else:
                status_str = "Cruzado 1 a 1" if s["conc"] > 0 else ("Falta Karing" if s["pb"] > 0 else "Falta Banco")
                disp_str = f"${a:,.2f}  |  Monto Único  ({status_str})"
            amt_options.append(disp_str)
            amt_map[disp_str] = a

        # Controles
        c_sel1, c_sel2, c_sel3 = st.columns([2.5, 1.5, 1.2])
        with c_sel1:
            selected_option = st.selectbox(
                "Seleccionar Monto Frecuente o con Discrepancias:",
                options=amt_options,
                key="sel_rep_amt"
            )
        with c_sel2:
            custom_input_amt = st.number_input("O digitar monto exacto a auditar ($):", min_value=0.0, step=1000.0, value=0.0, key="custom_rep_amt")
        with c_sel3:
            all_ctas_rep = ["TODAS"] + sorted(list(set([c["cuenta_nombre"] for c in all_conc_list] + [pb["cuenta_nombre"] for pb in all_pb_list])))
            sel_cta_rep = st.selectbox("Filtrar por Cuenta:", all_ctas_rep, key="sel_cta_rep")

        if custom_input_amt > 0:
            target_amt = round(custom_input_amt, 2)
        else:
            target_amt = amt_map.get(selected_option, 0.0)

        # Filtrar registros para el target_amt y cuenta
        cur_conc = [c for c in all_conc_list if round(c["valor_conciliado"], 2) == target_amt and (sel_cta_rep == "TODAS" or c["cuenta_nombre"] == sel_cta_rep)]
        cur_pb = [pb for pb in all_pb_list if round(abs(pb["monto"]), 2) == target_amt and (sel_cta_rep == "TODAS" or pb["cuenta_nombre"] == sel_cta_rep)]
        cur_pk = [pk for pk in all_pk_list if round(pk["debito"] if pk["debito"] > 0 else pk["credito"], 2) == target_amt and (sel_cta_rep == "TODAS" or pk["cuenta_nombre"] == sel_cta_rep)]

        tot_banco_cur = len(cur_conc) + len(cur_pb)
        tot_karing_cur = len(cur_conc) + len(cur_pk)

        # Métricas del monto
        st.markdown(f"#### 📊 Balance para el Monto: `${target_amt:,.2f}`")
        m_col1, m_col2, m_col3, m_col4 = st.columns(4)
        with m_col1:
            st.metric("Total en Extracto Bancario", f"{tot_banco_cur} transacciones", f"${tot_banco_cur * target_amt:,.2f}")
        with m_col2:
            st.metric("Total en Libros Karing", f"{tot_karing_cur} registros", f"${tot_karing_cur * target_amt:,.2f}")
        with m_col3:
            st.metric("✅ Machearon 1 a 1", f"{len(cur_conc)} pareos", f"${len(cur_conc) * target_amt:,.2f}")
        with m_col4:
            st.metric("❌ Pendientes de Cuadrar", f"{len(cur_pb) + len(cur_pk)} movimientos", f"${(len(cur_pb) + len(cur_pk)) * target_amt:,.2f}")

        # SECCIÓN 1: LOS QUE NO MACHEARON
        st.markdown("---")
        st.markdown("### ❌ 1. ¿Cuáles NO machearon? (Partidas Huérfanas de este Monto)")
        
        if not cur_pb and not cur_pk:
            st.success(f"🎉 **¡Perfecto!** Todos los movimientos de este monto (${target_amt:,.2f}) se conciliaron al 100%. No hay partidas pendientes.")
        else:
            if cur_pb:
                st.markdown(f"##### 🏦 Extracto Bancario: {len(cur_pb)} movimiento(s) que NO se cruzaron con Karing:")
                st.caption(f"Entraron/salieron del banco pero NO se encontró recibo/egreso equivalente en Karing:")
                df_cur_pb = pd.DataFrame(cur_pb)
                if "documento" not in df_cur_pb.columns:
                    df_cur_pb["documento"] = ""
                if "contexto_monto" not in df_cur_pb.columns:
                    df_cur_pb["contexto_monto"] = "Monto único"
                df_cur_pb["ref_banco"] = df_cur_pb["documento"].fillna("").astype(str).str.strip().replace({"": "(Sin ref. en extracto)", "None": "(Sin ref. en extracto)", "nan": "(Sin ref. en extracto)"})
                df_cur_pb_disp = df_cur_pb[[
                    "cuenta_nombre", "fecha", "ref_banco", "descripcion", "contexto_monto"
                ]].rename(columns={
                    "cuenta_nombre": "Cuenta Bancaria",
                    "fecha": "Fecha en Extracto",
                    "ref_banco": "No. Referencia Banco",
                    "descripcion": "Descripción en Extracto",
                    "contexto_monto": "Situación de Repetición"
                })
                st.dataframe(df_cur_pb_disp, width="stretch", hide_index=True)

            if cur_pk:
                st.markdown(f"##### 📖 Libros Karing: {len(cur_pk)} registro(s) que NO se reflejaron en el Banco:")
                st.caption(f"Fueron contabilizados en Karing pero la plata nunca ingresó/salió del extracto:")
                df_cur_pk = pd.DataFrame(cur_pk)
                for c, d in [("documento", ""), ("tipo_documento", "RC"), ("contexto_monto", "Monto único")]:
                    if c not in df_cur_pk.columns:
                        df_cur_pk[c] = d
                df_cur_pk_disp = df_cur_pk[[
                    "cuenta_nombre", "fecha", "documento", "tipo_documento", "detalle", "contexto_monto"
                ]].rename(columns={
                    "cuenta_nombre": "Cuenta en Karing",
                    "fecha": "Fecha Contable",
                    "documento": "No. Recibo / Doc",
                    "tipo_documento": "Tipo Doc",
                    "detalle": "Tercero / Concepto Karing",
                    "contexto_monto": "Situación de Repetición"
                })
                st.dataframe(df_cur_pk_disp, width="stretch", hide_index=True)

        # SECCIÓN 2: LOS QUE SÍ MACHEARON (PAREO EXACTO)
        st.markdown("---")
        st.markdown("### ✅ 2. ¿Cuáles SÍ machearon exactamente? (Trazabilidad de Pareos 1 a 1)")
        if cur_conc:
            st.caption(f"Mostrando los **{len(cur_conc)}** emparejamientos exactos realizados por cercanía de fecha:")
            df_cur_conc = pd.DataFrame(cur_conc)
            if "id_pareo" not in df_cur_conc.columns:
                df_cur_conc["id_pareo"] = [f"PAR-{i+1:04d}" for i in range(len(df_cur_conc))]
            df_cur_conc["documento_banco"] = df_cur_conc.get("documento_banco", pd.Series([""]*len(df_cur_conc))).fillna("").astype(str).str.strip().replace({"": "(Sin ref)", "None": "(Sin ref)", "nan": "(Sin ref)"})
            df_cur_conc["tipo_documento_karing"] = df_cur_conc.get("tipo_documento_karing", pd.Series(["RC"]*len(df_cur_conc))).fillna("RC")

            df_cur_conc_disp = df_cur_conc[[
                "id_pareo", "cuenta_nombre", "fecha_banco", "documento_banco", "descripcion_banco",
                "fecha_karing", "documento_karing", "tipo_documento_karing", "detalle_karing", "diferencia_dias"
            ]].rename(columns={
                "id_pareo": "ID Pareo",
                "cuenta_nombre": "Cuenta",
                "fecha_banco": "Fecha Banco",
                "documento_banco": "Ref. Banco",
                "descripcion_banco": "Descripción en Extracto",
                "fecha_karing": "Fecha Karing",
                "documento_karing": "Doc Karing",
                "tipo_documento_karing": "Tipo Doc",
                "detalle_karing": "Detalle en Karing",
                "diferencia_dias": "Dif. Días"
            })
            st.dataframe(df_cur_conc_disp, width="stretch", hide_index=True)

            # Botón de Descarga de este informe de auditoría específico
            excel_audit_bytes = safe_export_dataframe_to_excel_bytes(
                df_cur_conc_disp,
                title=f"Auditoría Pareos para Monto ${target_amt:,.2f} - {selected_month}",
                sheet_name="Pareos_Exitosos"
            )
            col_ad1, col_ad2 = st.columns([3, 1])
            with col_ad2:
                st.download_button(
                    label=f"📥 Descargar Pareos de ${target_amt:,.0f} en Excel (.xlsx)",
                    data=excel_audit_bytes,
                    file_name=f"Auditoria_Pareos_Monto_{int(target_amt)}_{selected_month.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch"
                )
        else:
            st.info(f"Ningún movimiento de este monto se ha conciliado aún.")

# =============================================================
# VISTA 8: CARGAR NUEVOS ARCHIVOS / MESES
# =============================================================
elif st.session_state["active_tab"] == MENU_OPTIONS[7]:
    st.markdown("""
    <div class="info-banner">
        <h4 style="margin:0 0 6px 0; color:#1e293b;">📤 Cargar Extractos y Libros Auxiliares para Nuevos Meses</h4>
        Sube los archivos correspondientes a meses futuros (ej. Septiembre 2026). El sistema los organizará automáticamente en la carpeta correspondiente.
    </div>
    """, unsafe_allow_html=True)

    col_u1, col_u2, col_u3 = st.columns(3)
    with col_u1:
        nuevo_mes = st.text_input("Nombre del Periodo (ej. SEPTIEMBRE 2026):", value="SEPTIEMBRE 2026")
    with col_u2:
        banco_destino = st.selectbox("Entidad Financiera:", options=["BANCOLOMBIA", "BBVA", "BOGOTA", "DAVIVIENDA", "BOLD"])
    with col_u3:
        cuenta_destino = st.selectbox(
            "Cuenta Destino:",
            options=[k for k, v in ACCOUNTS_CATALOG.items() if v.bank_name == banco_destino]
        )

    uploaded_files = st.file_uploader(
        "Arrastra los archivos aquí (Extracto PDF o Libro Auxiliar Excel):",
        accept_multiple_files=True,
        type=["pdf", "xls", "xlsx"]
    )

    if st.button("Guardar Archivos en el Sistema", width="stretch"):
        if uploaded_files:
            cfg = ACCOUNTS_CATALOG[cuenta_destino]
            mes_str = nuevo_mes.strip().upper()
            if cfg.subfolder_name:
                target_folder = os.path.join(ROOT_DIR, cfg.folder_name, mes_str, cfg.subfolder_name)
            else:
                target_folder = os.path.join(ROOT_DIR, cfg.folder_name, mes_str)
            os.makedirs(target_folder, exist_ok=True)

            for up_file in uploaded_files:
                dest_path = os.path.join(target_folder, up_file.name)
                with open(dest_path, "wb") as f_out:
                    f_out.write(up_file.getbuffer())
                st.success(f"Archivo guardado correctamente: `{dest_path}`")
            st.rerun()
        else:
            st.warning("Selecciona al menos un archivo para cargar.")
