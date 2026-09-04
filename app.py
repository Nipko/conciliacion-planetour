"""
Panel de Control Ejecutivo y Tablero de Conciliación Bancaria - Planetour SAS.
Interfaz moderna, intuitiva, con navegación interactiva y descarga individual de Excel por sección.
"""

import streamlit as st
import pandas as pd
import os
import io
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


# -------------------------------------------------------------
# GESTIÓN DE ESTADO DE NAVEGACIÓN (DRILL-DOWN INTERACTIVO)
# -------------------------------------------------------------
MENU_OPTIONS = [
    "📊 1. Tablero Ejecutivo de Control",
    "🚨 2. Lo que está MAL (Cruces Intercuentas)",
    "❓ 3. Lo que FALTA por Conciliar (Pendientes)",
    "💸 4. Gastos e Impuestos Bancarios (4x1000)",
    "✅ 5. Lo que está BIEN (Conciliados 1 a 1)",
    "🔍 6. Buscador y Auditoría de Montos Repetidos",
    "📤 7. Cargar Nuevos Archivos / Meses"
]

if "active_tab" not in st.session_state or st.session_state.get("active_tab") == "🔍 6. Buscador Universal de Movimientos":
    st.session_state["active_tab"] = MENU_OPTIONS[5] if st.session_state.get("active_tab") == "🔍 6. Buscador Universal de Movimientos" else MENU_OPTIONS[0]

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
def run_reconciliation(month: str, tol: int, _cache_version: str = "v2.2"):
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

    # 4 TARJETAS ACCIONABLES CON BOTÓN CLICKABLE
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown(f"""
        <div class="action-card" style="border-top: 4px solid #10b981;">
            <div class="action-card-title" style="color: #059669;">✅ Conciliados (1 a 1)</div>
            <div class="action-card-val">{glob['total_conciliados']:,} <span style="font-size:15px; color:#64748b;">partidas</span></div>
            <div class="action-card-desc">
                Monto: <b>${glob['monto_total_conciliado']:,.2f}</b><br>
                Coincidencia exacta de valor y fecha entre banco y Karing.
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("👉 Ver Movimientos Conciliados", key="btn_go_conc", width="stretch"):
            navigate_to(MENU_OPTIONS[4])

    with c2:
        st.markdown(f"""
        <div class="action-card" style="border-top: 4px solid #f59e0b;">
            <div class="action-card-title" style="color: #d97706;">🚨 Lo que está MAL (Intercuentas)</div>
            <div class="action-card-val">{glob['total_cruces_intercuentas']:,} <span style="font-size:15px; color:#64748b;">traslados</span></div>
            <div class="action-card-desc">
                Monto: <b>${glob['monto_total_cruces']:,.2f}</b><br>
                Dinero que llegó a un banco pero se asentó en otra cuenta en Karing.
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("👉 Corregir Cuentas Erróneas", key="btn_go_cross", width="stretch"):
            navigate_to(MENU_OPTIONS[1])

    with c3:
        st.markdown(f"""
        <div class="action-card" style="border-top: 4px solid #8b5cf6;">
            <div class="action-card-title" style="color: #7c3aed;">💸 Gastos e Impuestos Banco</div>
            <div class="action-card-val">${glob['total_gastos_bancarios']:,.0f}</div>
            <div class="action-card-desc">
                GMF 4x1000: <b>${glob['total_gmf']:,.0f}</b> | Comisiones: <b>${glob['total_comisiones']:,.0f}</b><br>
                Cobros automáticos del banco por registrar en Karing.
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("👉 Ver Gastos y Asientos", key="btn_go_tax", width="stretch"):
            navigate_to(MENU_OPTIONS[3])

    with c4:
        total_faltantes = glob['total_pendientes_banco'] + glob['total_pendientes_karing']
        st.markdown(f"""
        <div class="action-card" style="border-top: 4px solid #ef4444;">
            <div class="action-card-title" style="color: #dc2626;">❓ Lo que FALTA por Conciliar</div>
            <div class="action-card-val">{total_faltantes:,} <span style="font-size:15px; color:#64748b;">partidas</span></div>
            <div class="action-card-desc">
                En Banco sin Karing: <b>{glob['total_pendientes_banco']:,}</b> (${glob['monto_total_pendientes_banco']:,.0f})<br>
                En Karing sin Banco: <b>{glob['total_pendientes_karing']:,}</b> (${glob['monto_total_pendientes_karing']:,.0f})
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
        Movimientos automáticos cobrados o abonados por los bancos durante el mes. 
        Incluye el cálculo consolidado y la propuesta del asiento contable para asentar en Karing.
    </div>
    """, unsafe_allow_html=True)

    all_exp = []
    for acc_k, acc_r in reconciliation["cuentas"].items():
        all_exp.extend(acc_r.get("gastos_impuestos", []))

    if all_exp:
        df_exp = pd.DataFrame(all_exp)

        # Métricas de Conceptos
        m_g1, m_g2, m_g3, m_g4 = st.columns(4)
        with m_g1:
            st.metric("GMF (4x1000) Total", f"${glob['total_gmf']:,.2f}")
        with m_g2:
            st.metric("Comisiones ACH / Portales", f"${glob['total_comisiones']:,.2f}")
        with m_g3:
            st.metric("Intereses a Favor (Rendimientos)", f"${glob['total_intereses']:,.2f}")
        with m_g4:
            st.metric("Total Gastos e Impuestos", f"${glob['total_gastos_bancarios']:,.2f}")

        # Plantilla de Asiento Contable
        st.markdown("#### 📝 Asiento Contable Modelo Sugerido para Contabilizar en Karing:")
        st.code(f"""
-- ASIENTO DE AJUSTE BANCARIO - PERIODO {selected_month}
---------------------------------------------------------------------------------------------
CUENTA CONTABLE                    CONCEPTO                      DÉBITO ($)       CRÉDITO ($)
---------------------------------------------------------------------------------------------
511595 - GMF 4x1000                Impuesto GMF del mes          ${glob['total_gmf']:,.2f}
530515 - Comisiones Bancarias      Comisiones ACH / Efecty/Portal ${glob['total_comisiones']:,.2f}
11XXXX - Bancos (Varias Cuentas)   Salida por gastos bancarios                     ${glob['total_gastos_bancarios']:,.2f}
---------------------------------------------------------------------------------------------
11XXXX - Bancos Cuentas Ahorro     Entrada por rendimientos      ${glob['total_intereses']:,.2f}
421005 - Ingresos Financieros      Intereses abonados mes                         ${glob['total_intereses']:,.2f}
---------------------------------------------------------------------------------------------
        """, language="sql")

        st.markdown("#### Detalle Individual de Cada Cargo:")
        df_exp_disp = df_exp.copy()
        df_exp_disp["documento"] = df_exp_disp["documento"].fillna("").astype(str).str.strip().replace({"": "(Sin ref. en extracto)", "None": "(Sin ref. en extracto)", "nan": "(Sin ref. en extracto)"})
        
        df_exp_disp = df_exp_disp[[
            "banco", "cuenta_banco", "fecha", "monto", "tipo", "documento", "sugerencia_contable", "descripcion"
        ]].rename(columns={
            "banco": "Banco",
            "cuenta_banco": "Cuenta",
            "fecha": "Fecha",
            "monto": "Monto ($)",
            "tipo": "Tipo de Partida",
            "documento": "No. Referencia Banco",
            "sugerencia_contable": "Asiento Contable Sugerido",
            "descripcion": "Descripción en Extracto"
        })

        # Botón de Descarga Excel
        excel_exp_bytes = safe_export_dataframe_to_excel_bytes(
            df_exp_disp,
            title=f"Gastos e Impuestos Bancarios - {selected_month}",
            sheet_name="Gastos_Impuestos"
        )
        col_ex1, col_ex2 = st.columns([3, 1])
        with col_ex2:
            st.download_button(
                label="📥 Descargar Gastos e Impuestos en Excel (.xlsx)",
                data=excel_exp_bytes,
                file_name=f"Gastos_Impuestos_{selected_month.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch"
            )

        st.dataframe(
            df_exp_disp,
            column_config={
                "Banco": st.column_config.TextColumn("Banco", width=120),
                "Cuenta": st.column_config.TextColumn("Cuenta", width=150),
                "Fecha": st.column_config.TextColumn("Fecha", width=105),
                "Monto ($)": st.column_config.NumberColumn("Monto ($)", format="$ %,.2f", width=130),
                "Tipo de Partida": st.column_config.TextColumn("Tipo", width=120),
                "No. Referencia Banco": st.column_config.TextColumn("Ref. Banco", width=140),
                "Asiento Contable Sugerido": st.column_config.TextColumn("Asiento Contable Sugerido", width=320),
                "Descripción en Extracto": st.column_config.TextColumn("Descripción en Extracto", width=350)
            },
            width="stretch",
            hide_index=True
        )

# =============================================================
# VISTA 5: LO QUE ESTÁ BIEN (CONCILIADOS 1 A 1)
# =============================================================
elif st.session_state["active_tab"] == MENU_OPTIONS[4]:
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
# VISTA 6: BUSCADOR UNIVERSAL
# =============================================================
elif st.session_state["active_tab"] == MENU_OPTIONS[5]:
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
# VISTA 7: CARGAR NUEVOS ARCHIVOS / MESES
# =============================================================
elif st.session_state["active_tab"] == MENU_OPTIONS[6]:
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
