"""
Parser para reportes mensuales de transacciones de Bold en Excel.
"""

import pandas as pd
from datetime import datetime
from typing import Dict, Any, List, Optional
import os

class BoldParser:
    """
    Extrae transacciones de datáfonos/links de pago de Bold, deducciones, retenciones y depósito neto.
    """

    @staticmethod
    def parse_file(file_path: str) -> Dict[str, Any]:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"No se encontró el archivo: {file_path}")

        # Encontrar la fila de cabecera en las primeras 5 filas
        df_raw = pd.read_excel(file_path, header=None, nrows=6, engine="openpyxl")
        header_idx = 2
        for idx, r in df_raw.iterrows():
            row_str = " ".join([str(v).lower() for v in r.values if pd.notna(v)])
            if "transacci" in row_str and "fecha" in row_str and "total" in row_str:
                header_idx = idx
                break

        df = pd.read_excel(file_path, header=header_idx, engine="openpyxl")

        # Normalizar columnas para buscar por patrones
        col_map = {}
        for c in df.columns:
            clean = str(c).strip().lower()
            clean = clean.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
            col_map[clean] = c

        def get_col(*patterns):
            for p in patterns:
                for clean, orig in col_map.items():
                    if p in clean:
                        return orig
            return None

        c_id = get_col("id transaccion")
        c_fecha = get_col("fecha")
        c_estado = get_col("estado", "operacion")
        c_total = get_col("valor total")
        c_retefte = get_col("rete fuente")
        c_reteiva = get_col("rete iva")
        c_reteica = get_col("rete ica")
        c_deduccion = get_col("total deduccion")
        c_deposito = get_col("deposito en cuenta", "deposito en saldo")
        c_datafono = get_col("nombre del datafono")
        c_metodo = get_col("metodo de pago")

        if not c_total or not c_fecha:
            raise ValueError(f"No se encontraron las columnas esenciales en el reporte de Bold: {file_path}")

        # Eliminar filas vacías
        df = df.dropna(subset=[c_fecha, c_total]).copy()

        df["fecha_dt"] = pd.to_datetime(df[c_fecha], errors="coerce")
        df["valor_total"] = pd.to_numeric(df[c_total], errors="coerce").fillna(0.0).round(2)
        df["rete_fuente"] = pd.to_numeric(df[c_retefte], errors="coerce").fillna(0.0).round(2) if c_retefte else 0.0
        df["rete_iva"] = pd.to_numeric(df[c_reteiva], errors="coerce").fillna(0.0).round(2) if c_reteiva else 0.0
        df["rete_ica"] = pd.to_numeric(df[c_reteica], errors="coerce").fillna(0.0).round(2) if c_reteica else 0.0
        df["total_deduccion"] = pd.to_numeric(df[c_deduccion], errors="coerce").fillna(0.0).round(2) if c_deduccion else 0.0
        df["deposito_neto"] = pd.to_numeric(df[c_deposito], errors="coerce").fillna(0.0).round(2) if c_deposito else 0.0

        transactions = []
        for _, row in df.iterrows():
            f_dt = row["fecha_dt"]
            t_id = str(row[c_id]) if c_id and pd.notna(row.get(c_id)) else ""
            datafono = str(row[c_datafono]) if c_datafono and pd.notna(row.get(c_datafono)) else "BOLD"
            val_bruto = float(row["valor_total"])
            
            transactions.append({
                "fecha": f_dt,
                "fecha_str": f_dt.strftime("%Y-%m-%d") if pd.notna(f_dt) else "",
                "documento": t_id,
                "descripcion": f"VENTA BOLD {datafono}".strip(),
                "monto": val_bruto,
                "es_abono": True,
                "es_cargo": False,
                "saldo": None,
                "valor_bruto": val_bruto,
                "rete_fuente": float(row["rete_fuente"]),
                "rete_iva": float(row["rete_iva"]),
                "rete_ica": float(row["rete_ica"]),
                "comisiones_bold": float(row["total_deduccion"]),
                "deposito_neto": float(row["deposito_neto"]),
                "nombre_datafono": datafono,
                "metodo_pago": str(row[c_metodo]) if c_metodo and pd.notna(row.get(c_metodo)) else "",
                "categoria": "OPERATIVA",
                "banco": "BOLD",
                "cuenta_banco": "1700-1210-7997",
                "archivo_origen": os.path.basename(file_path)
            })

        tot_bruto = float(df["valor_total"].sum())
        tot_rf = float(df["rete_fuente"].sum()) if isinstance(df["rete_fuente"], pd.Series) else 0.0
        tot_riva = float(df["rete_iva"].sum()) if isinstance(df["rete_iva"], pd.Series) else 0.0
        tot_rica = float(df["rete_ica"].sum()) if isinstance(df["rete_ica"], pd.Series) else 0.0
        tot_ded = float(df["total_deduccion"].sum()) if isinstance(df["total_deduccion"], pd.Series) else 0.0
        tot_neto = float(df["deposito_neto"].sum()) if isinstance(df["deposito_neto"], pd.Series) else 0.0

        return {
            "banco": "BOLD",
            "archivo": file_path,
            "nombre_archivo": os.path.basename(file_path),
            "cuenta_numero": "1700-1210-7997",
            "saldo_anterior": 0.0,
            "total_abonos": round(tot_bruto, 2),
            "total_cargos": round(tot_ded + tot_rf, 2),
            "saldo_final": round(tot_neto, 2),
            "total_bruto": round(tot_bruto, 2),
            "total_retefuente": round(tot_rf, 2),
            "total_reteiva": round(tot_riva, 2),
            "total_reteica": round(tot_rica, 2),
            "total_comisiones": round(tot_ded, 2),
            "total_deposito_neto": round(tot_neto, 2),
            "num_transacciones": len(transactions),
            "transacciones": transactions
        }
