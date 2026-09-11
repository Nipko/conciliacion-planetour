"""
Generador de Informes Excel Profesionales de Conciliación Bancaria para Planetour SAS.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
from typing import Dict, Any, List
import os
import io
import pandas as pd

class ExcelReportGenerator:
    """
    Crea un libro Excel corporativo, formateado y con fórmulas
    para la presentación formal de la conciliación bancaria mensual.
    """

    def __init__(self):
        # Paleta de colores corporativos
        self.color_header_bg = "1B365D"      # Azul marino oscuro
        self.color_header_fg = "FFFFFF"      # Blanco
        self.color_accent_bg = "4A90E2"      # Azul medio
        self.color_alert_bg = "FFF2D6"       # Amarillo pastel suave
        self.color_alert_fg = "8A6D3B"       # Marrón dorado
        self.color_error_bg = "F8D7DA"       # Rojo pastel suave
        self.color_error_fg = "721C24"       # Rojo oscuro
        self.color_success_bg = "D4EDDA"     # Verde pastel suave
        self.color_zebra = "F9FAFB"          # Gris muy tenue

        # Fuentes
        self.font_title = Font(name="Calibri", size=16, bold=True, color="1B365D")
        self.font_subtitle = Font(name="Calibri", size=11, italic=True, color="555555")
        self.font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        self.font_data = Font(name="Calibri", size=10)
        self.font_bold = Font(name="Calibri", size=10, bold=True)
        self.font_alert = Font(name="Calibri", size=10, bold=True, color="8A6D3B")

        # Rellenos
        self.fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
        self.fill_subtotal = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid")
        self.fill_alert = PatternFill(start_color="FFF2D6", end_color="FFF2D6", fill_type="solid")
        self.fill_zebra = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")

        # Bordes
        thin = Side(border_style="thin", color="D3D3D3")
        self.border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)
        self.border_top_thick = Border(top=Side(border_style="medium", color="1B365D"))

        # Formatos de número
        self.fmt_currency = "$#,##0.00"
        self.fmt_number = "#,##0"
        self.fmt_date = "YYYY-MM-DD"

    def generate_report_workbook(self, reconciliation_data: Dict[str, Any]) -> openpyxl.Workbook:
        """
        Construye el Workbook con todas las hojas de conciliación.
        """
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        # 1. Hoja: Resumen Ejecutivo
        self._build_executive_summary_sheet(wb, reconciliation_data)

        # 2. Hoja: Cruces Intercuentas (Anomalías clave)
        self._build_cross_account_sheet(wb, reconciliation_data)

        # 3. Hoja: Gastos e Impuestos Bancarios (GMF, Comisiones, IVA)
        self._build_bank_expenses_sheet(wb, reconciliation_data)

        # 4. Hoja: Intereses y Rendimientos Financieros
        self._build_interest_sheet(wb, reconciliation_data)

        # 5. Hoja: Conciliados 1-a-1
        self._build_matched_sheet(wb, reconciliation_data)

        # 5. Hoja: Pendientes en Banco
        self._build_pending_bank_sheet(wb, reconciliation_data)

        # 6. Hoja: Pendientes en Karing
        self._build_pending_karing_sheet(wb, reconciliation_data)

        return wb

    def generate_report_bytes(self, reconciliation_data: Dict[str, Any]) -> bytes:
        """
        Genera el informe Excel directamente en bytes en memoria.
        Inmune a bloqueos de archivos abiertos en Excel.
        """
        wb = self.generate_report_workbook(reconciliation_data)
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def generate_report(self, reconciliation_data: Dict[str, Any], output_path: str) -> str:
        """
        Genera y guarda el archivo Excel en disco, manejando posibles bloqueos si está abierto en Excel.
        """
        wb = self.generate_report_workbook(reconciliation_data)
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        try:
            wb.save(output_path)
            return output_path
        except PermissionError:
            base, ext = os.path.splitext(output_path)
            alt_path = f"{base}_{datetime.now().strftime('%H%M%S')}{ext}"
            wb.save(alt_path)
            return alt_path

    def export_dataframe_to_excel_bytes(self, df: pd.DataFrame, title: str, sheet_name: str = "Datos") -> bytes:
        """
        Exporta cualquier DataFrame a un Excel formateado profesionalmente en memoria.
        Ideal para descargas individuales por cada sección del dashboard.
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name[:31]
        ws.views.sheetView[0].showGridLines = True

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        # Título
        ws["A1"] = f"PLANETOUR S.A.S. — {title.upper()}"
        ws["A1"].font = self.font_title
        ws["A2"] = f"Fecha de Generación: {now_str}"
        ws["A2"].font = self.font_subtitle

        # Cabeceras de columna en fila 4
        row_idx = 4
        headers = list(df.columns)
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=str(h))
            cell.font = self.font_header
            cell.fill = self.fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.row_dimensions[row_idx].height = 25
        row_idx += 1

        # Filas de datos
        for _, r in df.iterrows():
            for col_idx, col_name in enumerate(headers, 1):
                val = r[col_name]
                cell = ws.cell(row=row_idx, column=col_idx)

                # Detección de tipos y formatos
                if pd.isna(val):
                    cell.value = ""
                elif isinstance(val, (int, float)):
                    cell.value = float(val)
                    col_lower = str(col_name).lower()
                    if any(k in col_lower for k in ["valor", "monto", "debito", "credito", "saldo", "gmf", "comision", "$"]):
                        cell.number_format = self.fmt_currency
                    else:
                        cell.number_format = self.fmt_number
                    cell.alignment = Alignment(horizontal="right")
                else:
                    cell.value = str(val)
                    cell.alignment = Alignment(horizontal="left")

                cell.border = self.border_cell
                cell.font = self.font_data
            row_idx += 1

        self._auto_fit_columns(ws)

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def _build_executive_summary_sheet(self, wb: openpyxl.Workbook, data: Dict[str, Any]):
        ws = wb.create_sheet(title="Resumen Ejecutivo")
        ws.views.sheetView[0].showGridLines = True

        mes = data.get("mes", "PERIODO")
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        # Título
        ws["A1"] = "PLANETOUR S.A.S. - INFORME MENSUAL DE CONCILIACIÓN BANCARIA"
        ws["A1"].font = self.font_title
        ws["A2"] = f"Periodo Auditado: {mes} | Fecha de Generación: {now_str}"
        ws["A2"].font = self.font_subtitle

        # Advertencias si existen
        row_idx = 4
        advertencias = data.get("advertencias", [])
        if advertencias:
            ws.cell(row=row_idx, column=1, value="⚠️ ADVERTENCIAS DE AUDITORÍA DETECTADAS:").font = self.font_alert
            row_idx += 1
            for adv in advertencias:
                c = ws.cell(row=row_idx, column=1, value=f"• {adv}")
                c.font = self.font_subtitle
                row_idx += 1
            row_idx += 1

        # Tabla de Resumen por Cuenta
        headers = [
            "Banco", "Cuenta Contable", "Nombre de Cuenta Karing", "No. Cuenta / Ref",
            "Saldo Banco ($)", "Saldo Karing ($)", "Diferencia Bruta ($)",
            "Conciliados (#)", "Cruce Intercuentas (#)", "Gastos/Impuestos ($)",
            "GMF 4x1000 ($)", "Comisiones ($)", "Intereses ($)",
            "Pendientes Banco (#)", "Pendientes Karing (#)"
        ]

        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=h)
            cell.font = self.font_header
            cell.fill = self.fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws.row_dimensions[row_idx].height = 28
        row_idx += 1

        cuentas = data.get("cuentas", {})
        start_data_row = row_idx

        for acc_key, acc_res in cuentas.items():
            cfg = acc_res["config"]
            m = acc_res["metrics"]

            ws.cell(row=row_idx, column=1, value=cfg.bank_name).alignment = Alignment(horizontal="left")
            ws.cell(row=row_idx, column=2, value=cfg.karing_code).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=3, value=cfg.karing_name).alignment = Alignment(horizontal="left")
            ws.cell(row=row_idx, column=4, value=str(cfg.account_number)).alignment = Alignment(horizontal="center")

            # Valores numéricos
            c_sb = ws.cell(row=row_idx, column=5, value=m["saldo_banco"])
            c_sb.number_format = self.fmt_currency
            
            c_sk = ws.cell(row=row_idx, column=6, value=m["saldo_karing"])
            c_sk.number_format = self.fmt_currency

            c_dif = ws.cell(row=row_idx, column=7, value=m["diferencia_bruta"])
            c_dif.number_format = self.fmt_currency

            ws.cell(row=row_idx, column=8, value=m["total_conciliados"]).alignment = Alignment(horizontal="center")
            
            c_ci = ws.cell(row=row_idx, column=9, value=m["num_cruces_intercuentas"])
            c_ci.alignment = Alignment(horizontal="center")
            if m["num_cruces_intercuentas"] > 0:
                c_ci.font = self.font_bold
                c_ci.fill = self.fill_alert

            c_gt = ws.cell(row=row_idx, column=10, value=m["total_gastos_bancarios"])
            c_gt.number_format = self.fmt_currency

            c_gmf = ws.cell(row=row_idx, column=11, value=m["total_gmf"])
            c_gmf.number_format = self.fmt_currency

            c_com = ws.cell(row=row_idx, column=12, value=m["total_comisiones"])
            c_com.number_format = self.fmt_currency

            c_int = ws.cell(row=row_idx, column=13, value=m["total_intereses"])
            c_int.number_format = self.fmt_currency

            ws.cell(row=row_idx, column=14, value=m["num_pendientes_banco"]).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=15, value=m["num_pendientes_karing"]).alignment = Alignment(horizontal="center")

            # Aplicar bordes
            for col_c in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col_c).border = self.border_cell

            row_idx += 1

        # Fila de Totales
        end_data_row = row_idx - 1
        ws.cell(row=row_idx, column=1, value="TOTAL CONSOLIDADO").font = self.font_bold
        for col_c in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_c)
            cell.fill = self.fill_subtotal
            cell.border = self.border_cell
            cell.font = self.font_bold

        # Fórmulas de total
        ws.cell(row=row_idx, column=5, value=f"=SUM(E{start_data_row}:E{end_data_row})").number_format = self.fmt_currency
        ws.cell(row=row_idx, column=6, value=f"=SUM(F{start_data_row}:F{end_data_row})").number_format = self.fmt_currency
        ws.cell(row=row_idx, column=7, value=f"=SUM(G{start_data_row}:G{end_data_row})").number_format = self.fmt_currency
        ws.cell(row=row_idx, column=8, value=f"=SUM(H{start_data_row}:H{end_data_row})").number_format = self.fmt_number
        ws.cell(row=row_idx, column=9, value=f"=SUM(I{start_data_row}:I{end_data_row})").number_format = self.fmt_number
        ws.cell(row=row_idx, column=10, value=f"=SUM(J{start_data_row}:J{end_data_row})").number_format = self.fmt_currency
        ws.cell(row=row_idx, column=11, value=f"=SUM(K{start_data_row}:K{end_data_row})").number_format = self.fmt_currency
        ws.cell(row=row_idx, column=12, value=f"=SUM(L{start_data_row}:L{end_data_row})").number_format = self.fmt_currency
        ws.cell(row=row_idx, column=13, value=f"=SUM(M{start_data_row}:M{end_data_row})").number_format = self.fmt_currency
        ws.cell(row=row_idx, column=14, value=f"=SUM(N{start_data_row}:N{end_data_row})").number_format = self.fmt_number
        ws.cell(row=row_idx, column=15, value=f"=SUM(O{start_data_row}:O{end_data_row})").number_format = self.fmt_number

        self._auto_fit_columns(ws)

    def _build_cross_account_sheet(self, wb: openpyxl.Workbook, data: Dict[str, Any]):
        ws = wb.create_sheet(title="Cruces Intercuentas (Alertas)")
        ws.views.sheetView[0].showGridLines = True

        ws["A1"] = "MOVIMIENTOS IMPUTADOS EN CUENTA EQUIVOCADA (CRUCES INTERCUENTAS)"
        ws["A1"].font = self.font_title
        ws["A2"] = "IMPORTANTE: Estas partidas SÍ existen en Karing y en el Banco, pero contabilidad las registró en un banco diferente al real (ej: registradas en BBVA pero entraron a Bancolombia). NO deben volverse a registrar; únicamente trasladar en Karing."
        ws["A2"].font = self.font_subtitle

        headers = [
            "Fecha Banco", "Banco Real (Extracto)", "Descripción en Extracto", "Sentido",
            "Valor ($)", "Cuenta Errónea en Karing", "Cód Karing", "Doc Karing",
            "Fecha Karing", "Detalle en Karing", "Acción Correctiva Sugerida"
        ]

        row_idx = 4
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=h)
            cell.font = self.font_header
            cell.fill = self.fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.row_dimensions[row_idx].height = 25
        row_idx += 1

        all_cross = []
        for acc_k, acc_r in data.get("cuentas", {}).items():
            all_cross.extend(acc_r.get("cruces_intercuentas", []))

        # Ordenar por fecha
        all_cross.sort(key=lambda x: x.get("fecha_banco", ""))

        for item in all_cross:
            ws.cell(row=row_idx, column=1, value=item.get("fecha_banco", "")).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=2, value=item.get("banco_real", "")).alignment = Alignment(horizontal="left")
            ws.cell(row=row_idx, column=3, value=item.get("descripcion_banco", "")).alignment = Alignment(horizontal="left")
            ws.cell(row=row_idx, column=4, value=item.get("sentido", "")).alignment = Alignment(horizontal="center")

            c_val = ws.cell(row=row_idx, column=5, value=item.get("valor", 0.0))
            c_val.number_format = self.fmt_currency

            ws.cell(row=row_idx, column=6, value=item.get("cuenta_registrada_karing", "")).alignment = Alignment(horizontal="left")
            ws.cell(row=row_idx, column=7, value=item.get("codigo_cuenta_karing", "")).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=8, value=str(item.get("documento_karing", ""))).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=9, value=item.get("fecha_karing", "")).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=10, value=item.get("detalle_karing", "")).alignment = Alignment(horizontal="left")
            
            c_act = ws.cell(row=row_idx, column=11, value=item.get("accion_recomendada", ""))
            c_act.font = self.font_bold
            c_act.fill = self.fill_alert

            for col_c in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col_c).border = self.border_cell

            row_idx += 1

        self._auto_fit_columns(ws)

    def _build_bank_expenses_sheet(self, wb: openpyxl.Workbook, data: Dict[str, Any]):
        ws = wb.create_sheet(title="Gastos e Impuestos Bancarios")
        ws.views.sheetView[0].showGridLines = True

        ws["A1"] = "GASTOS BANCARIOS E IMPUESTOS (GMF 4X1000, COMISIONES, RETENCIONES, IVA)"
        ws["A1"].font = self.font_title
        ws["A2"] = "Listado detallado de cobros e impuestos deducidos por las entidades financieras para su registro contable."
        ws["A2"].font = self.font_subtitle

        headers = [
            "Banco", "Cuenta Planetour", "Tipo de Partida", "Fecha",
            "Documento / Ref", "Descripción Banco", "Monto ($)", "Asiento Contable Sugerido"
        ]

        row_idx = 4
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=h)
            cell.font = self.font_header
            cell.fill = self.fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.row_dimensions[row_idx].height = 25
        row_idx += 1

        all_exp = []
        for acc_k, acc_r in data.get("cuentas", {}).items():
            all_exp.extend(acc_r.get("gastos_impuestos", []))

        # Filtrar solo gastos e impuestos (excluyendo abonos de intereses)
        gastos_solo = [item for item in all_exp if item.get("tipo") != "INTERESES" and "RENDIMIENTO" not in str(item.get("descripcion", "")).upper()]

        total_gastos = 0.0
        for item in gastos_solo:
            ws.cell(row=row_idx, column=1, value=item.get("banco", "")).alignment = Alignment(horizontal="left")
            ws.cell(row=row_idx, column=2, value=item.get("cuenta_banco", "")).alignment = Alignment(horizontal="left")
            ws.cell(row=row_idx, column=3, value=item.get("tipo", "")).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=4, value=item.get("fecha", "")).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=5, value=str(item.get("documento", ""))).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=6, value=item.get("descripcion", "")).alignment = Alignment(horizontal="left")

            val = item.get("monto", 0.0)
            total_gastos += abs(val)
            c_val = ws.cell(row=row_idx, column=7, value=val)
            c_val.number_format = self.fmt_currency

            ws.cell(row=row_idx, column=8, value=item.get("sugerencia_contable", "")).alignment = Alignment(horizontal="left")

            for col_c in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col_c).border = self.border_cell

            row_idx += 1

        self._auto_fit_columns(ws)

    def _build_interest_sheet(self, wb: openpyxl.Workbook, data: Dict[str, Any]):
        ws = wb.create_sheet(title="Intereses y Rendimientos")
        ws.views.sheetView[0].showGridLines = True

        ws["A1"] = "INTERESES Y RENDIMIENTOS FINANCIEROS ABONADOS POR BANCOS"
        ws["A1"].font = self.font_title
        ws["A2"] = "Abonos de intereses en cuentas de ahorro, inversiones virtuales y CDTs para causación en Karing (Cuenta 421005)."
        ws["A2"].font = self.font_subtitle

        headers = [
            "Banco", "Cuenta Planetour", "Fecha Abono",
            "Descripción en Extracto", "Valor Rendimiento ($)", "Asiento Contable Sugerido"
        ]

        row_idx = 4
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=h)
            cell.font = self.font_header
            cell.fill = self.fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.row_dimensions[row_idx].height = 25
        row_idx += 1

        all_exp = []
        for acc_k, acc_r in data.get("cuentas", {}).items():
            all_exp.extend(acc_r.get("gastos_impuestos", []))

        intereses = [item for item in all_exp if item.get("tipo") == "INTERESES" or "RENDIMIENTO" in str(item.get("descripcion", "")).upper()]

        total_int = 0.0
        for item in intereses:
            ws.cell(row=row_idx, column=1, value=item.get("banco", "")).alignment = Alignment(horizontal="left")
            ws.cell(row=row_idx, column=2, value=item.get("cuenta_banco", "")).alignment = Alignment(horizontal="left")
            ws.cell(row=row_idx, column=3, value=item.get("fecha", "")).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=4, value=item.get("descripcion", "")).alignment = Alignment(horizontal="left")

            val = abs(float(item.get("monto", 0.0)))
            total_int += val
            c_val = ws.cell(row=row_idx, column=5, value=val)
            c_val.number_format = self.fmt_currency

            ws.cell(row=row_idx, column=6, value=item.get("sugerencia_contable", "")).alignment = Alignment(horizontal="left")

            for col_c in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col_c).border = self.border_cell

            row_idx += 1

        if intereses:
            ws.cell(row=row_idx, column=1, value="TOTAL RENDIMIENTOS").font = self.font_bold
            ws.cell(row=row_idx, column=1).alignment = Alignment(horizontal="left")
            ws.cell(row=row_idx, column=1).fill = self.fill_subtotal
            for c in range(2, 5):
                ws.cell(row=row_idx, column=c).fill = self.fill_subtotal
            c_tot = ws.cell(row=row_idx, column=5, value=total_int)
            c_tot.font = self.font_bold
            c_tot.number_format = self.fmt_currency
            c_tot.fill = self.fill_subtotal
            ws.cell(row=row_idx, column=6).fill = self.fill_subtotal
            for col_c in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col_c).border = self.border_cell

        self._auto_fit_columns(ws)

    def _build_matched_sheet(self, wb: openpyxl.Workbook, data: Dict[str, Any]):
        ws = wb.create_sheet(title="Partidas Conciliadas (1-a-1)")
        ws.views.sheetView[0].showGridLines = True

        ws["A1"] = "MOVIMIENTOS CONCILIADOS EXITOSAMENTE"
        ws["A1"].font = self.font_title

        headers = [
            "ID Pareo", "Cuenta", "Sentido", "Valor Conciliado ($)", "Fecha Banco", "Ref. Banco",
            "Descripción Banco", "Fecha Karing", "Doc Karing", "Tipo Doc", "Detalle Karing", "Dif. Días"
        ]

        row_idx = 3
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=h)
            cell.font = self.font_header
            cell.fill = self.fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.row_dimensions[row_idx].height = 25
        row_idx += 1

        all_mat = []
        for acc_k, acc_r in data.get("cuentas", {}).items():
            all_mat.extend(acc_r.get("conciliados", []))

        for item in all_mat:
            ref_banco = str(item.get("documento_banco", "")).strip()
            if not ref_banco or ref_banco.lower() in ["none", "nan", ""]:
                ref_banco = "(Sin ref)"

            ws.cell(row=row_idx, column=1, value=item.get("id_pareo", "")).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=2, value=item.get("cuenta_karing", "")).alignment = Alignment(horizontal="left")
            ws.cell(row=row_idx, column=3, value=item.get("sentido", "")).alignment = Alignment(horizontal="center")

            c_val = ws.cell(row=row_idx, column=4, value=item.get("valor_conciliado", 0.0))
            c_val.number_format = self.fmt_currency

            ws.cell(row=row_idx, column=5, value=item.get("fecha_banco", "")).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=6, value=ref_banco).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=7, value=item.get("descripcion_banco", "")).alignment = Alignment(horizontal="left")
            ws.cell(row=row_idx, column=8, value=item.get("fecha_karing", "")).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=9, value=str(item.get("documento_karing", ""))).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=10, value=str(item.get("tipo_documento_karing", "RC"))).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=11, value=item.get("detalle_karing", "")).alignment = Alignment(horizontal="left")
            ws.cell(row=row_idx, column=12, value=item.get("diferencia_dias", 0)).alignment = Alignment(horizontal="center")

            for col_c in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col_c).border = self.border_cell

            row_idx += 1

        self._auto_fit_columns(ws)

    def _build_pending_bank_sheet(self, wb: openpyxl.Workbook, data: Dict[str, Any]):
        ws = wb.create_sheet(title="Pendientes en Banco")
        ws.views.sheetView[0].showGridLines = True

        ws["A1"] = "MOVIMIENTOS EN EXTRACTO BANCARIO PENDIENTES POR REGISTRAR EN KARING"
        ws["A1"].font = self.font_title
        ws["A2"] = "NOTA: Si un movimiento del extracto no aparece en esta lista, consulte la hoja 'Cruces Intercuentas'. Las partidas consignadas o transferidas que contabilidad registró en otra cuenta bancaria se detallan allí para evitar duplicidades."
        ws["A2"].font = self.font_subtitle

        headers = [
            "Cuenta Bancaria", "Fecha Extracto", "Monto ($)", "Tipo de Movimiento",
            "Descripción en Extracto", "No. Referencia Banco", "Qué Falta Hacer en Karing"
        ]

        row_idx = 4
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=h)
            cell.font = self.font_header
            cell.fill = self.fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.row_dimensions[row_idx].height = 25
        row_idx += 1

        for acc_k, acc_r in data.get("cuentas", {}).items():
            for item in acc_r.get("pendientes_banco", []):
                t_raw = item.get("tipo", "")
                t_humano = "Abono / Ingreso" if "ABONO" in t_raw else "Cargo / Egreso"
                ref_doc = str(item.get("documento", "")).strip()
                if not ref_doc or ref_doc.lower() in ["none", "nan", ""]:
                    ref_doc = "(Sin ref. en extracto)"
                accion_karing = "Elaborar Recibo de Caja (RC)" if "Abono" in t_humano else "Elaborar Comprobante de Egreso (CE)"

                ws.cell(row=row_idx, column=1, value=item.get("cuenta", "")).alignment = Alignment(horizontal="left")
                ws.cell(row=row_idx, column=2, value=item.get("fecha", "")).alignment = Alignment(horizontal="center")

                c_val = ws.cell(row=row_idx, column=3, value=item.get("monto", 0.0))
                c_val.number_format = self.fmt_currency

                ws.cell(row=row_idx, column=4, value=t_humano).alignment = Alignment(horizontal="center")
                ws.cell(row=row_idx, column=5, value=item.get("descripcion", "")).alignment = Alignment(horizontal="left")
                ws.cell(row=row_idx, column=6, value=ref_doc).alignment = Alignment(horizontal="center")
                ws.cell(row=row_idx, column=7, value=accion_karing).alignment = Alignment(horizontal="left")

                for col_c in range(1, len(headers) + 1):
                    ws.cell(row=row_idx, column=col_c).border = self.border_cell
                row_idx += 1

        self._auto_fit_columns(ws)

    def _build_pending_karing_sheet(self, wb: openpyxl.Workbook, data: Dict[str, Any]):
        ws = wb.create_sheet(title="Pendientes en Karing")
        ws.views.sheetView[0].showGridLines = True

        ws["A1"] = "REGISTROS EN KARING PENDIENTES EN BANCO (CHEQUES O RECIBOS EN TRÁNSITO)"
        ws["A1"].font = self.font_title
        ws["A2"] = "NOTA: Si un recibo de caja o comprobante de egreso no aparece aquí, verifique la hoja 'Cruces Intercuentas'. Los registros imputados en un banco diferente al real se encuentran clasificados allí como traslados contables."
        ws["A2"].font = self.font_subtitle

        headers = [
            "Cuenta en Karing", "Fecha Contable", "No. Recibo / Doc Karing", "Tipo Doc (RC/CE)",
            "Tercero / Concepto en Karing", "Débito ($)", "Crédito ($)", "Estado de la Partida", "Acción Recomendada"
        ]

        row_idx = 4
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=h)
            cell.font = self.font_header
            cell.fill = self.fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.row_dimensions[row_idx].height = 25
        row_idx += 1

        for acc_k, acc_r in data.get("cuentas", {}).items():
            for item in acc_r.get("pendientes_karing", []):
                t_raw = item.get("tipo", "")
                is_ingreso = "DEBITO" in t_raw
                estado_humano = "Ingreso en Karing (No en Extracto)" if is_ingreso else "Egreso en Karing (No en Extracto)"
                accion = "Verificar si la consignación no ha ingresado al banco o anular recibo si fue devuelto" if is_ingreso else "Verificar cheque girado pendiente de cobro o débito diferido"

                ws.cell(row=row_idx, column=1, value=item.get("cuenta", "")).alignment = Alignment(horizontal="left")
                ws.cell(row=row_idx, column=2, value=item.get("fecha", "")).alignment = Alignment(horizontal="center")
                ws.cell(row=row_idx, column=3, value=str(item.get("documento", ""))).alignment = Alignment(horizontal="center")
                ws.cell(row=row_idx, column=4, value=str(item.get("tipo_documento", ""))).alignment = Alignment(horizontal="center")
                ws.cell(row=row_idx, column=5, value=item.get("detalle", "")).alignment = Alignment(horizontal="left")

                c_deb = ws.cell(row=row_idx, column=6, value=item.get("debito", 0.0))
                c_deb.number_format = self.fmt_currency

                c_cred = ws.cell(row=row_idx, column=7, value=item.get("credito", 0.0))
                c_cred.number_format = self.fmt_currency

                ws.cell(row=row_idx, column=8, value=estado_humano).alignment = Alignment(horizontal="center")
                ws.cell(row=row_idx, column=9, value=accion).alignment = Alignment(horizontal="left")

                for col_c in range(1, len(headers) + 1):
                    ws.cell(row=row_idx, column=col_c).border = self.border_cell
                row_idx += 1

        self._auto_fit_columns(ws)

    def _auto_fit_columns(self, ws: openpyxl.worksheet.worksheet.Worksheet):
        """Ajusta automáticamente el ancho de las columnas."""
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = cell.value
                if val is not None:
                    s_val = str(val)
                    if not s_val.startswith("="):
                        max_len = max(max_len, len(s_val))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
