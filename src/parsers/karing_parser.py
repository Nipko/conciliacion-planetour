"""
Parser para los libros auxiliares contables de Karing (formato Excel / OpenXML).
"""

import pandas as pd
from typing import Dict, Any, List, Optional
import os

class KaringParser:
    """
    Lee y estandariza los libros auxiliares generados por el sistema contable Karing.
    """

    @staticmethod
    def parse_file(file_path: str) -> Dict[str, Any]:
        """
        Lee un archivo de libro auxiliar de Karing y retorna la metadata y las transacciones.
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"No se encontró el archivo: {file_path}")

        # Leer las primeras 10 filas para encontrar dinámicamente la fila de cabeceras
        df_raw = pd.read_excel(file_path, header=None, nrows=10, engine="openpyxl")
        header_row_idx = None
        for idx, row in df_raw.iterrows():
            row_str = " ".join([str(v).lower() for v in row.values if pd.notna(v)])
            if "documento" in row_str and "debito" in row_str and "cuenta" in row_str:
                header_row_idx = idx
                break

        if header_row_idx is None:
            # Fallback histórico estándar de Karing: fila índice 4 (fila 5 de Excel)
            header_row_idx = 4

        df = pd.read_excel(file_path, skiprows=header_row_idx, engine="openpyxl")
        
        # Limpieza de nombres de columnas
        df.columns = [str(c).strip().lower() for c in df.columns]

        # Mapeo y validación de columnas esenciales
        req_cols = ["fecha", "documento", "cuenta", "detalle", "debito", "credito"]
        for col in req_cols:
            if col not in df.columns:
                raise ValueError(f"Columna obligatoria '{col}' no encontrada en {file_path}. Columnas: {df.columns.tolist()}")

        # Eliminar filas vacías de datos
        df = df.dropna(subset=["fecha", "documento", "cuenta"]).copy()

        # Conversión de tipos
        df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
        df["documento"] = df["documento"].astype(str).str.replace(r"\.0$", "", regex=True)
        df["cuenta"] = pd.to_numeric(df["cuenta"], errors="coerce").fillna(0).astype(int)
        df["debito"] = pd.to_numeric(df["debito"], errors="coerce").fillna(0.0).round(2)
        df["credito"] = pd.to_numeric(df["credito"], errors="coerce").fillna(0.0).round(2)
        df["detalle"] = df["detalle"].fillna("").astype(str).str.strip()

        # Saldo anterior (si existe en la fila)
        saldo_anterior = 0.0
        if "saldo_anterior" in df.columns:
            val = df["saldo_anterior"].dropna()
            if len(val) > 0:
                saldo_anterior = float(val.iloc[0])

        cuenta_code = int(df["cuenta"].iloc[0]) if len(df) > 0 else 0
        cuenta_desc = str(df["descripcion_cuenta"].iloc[0]) if "descripcion_cuenta" in df.columns and len(df) > 0 else ""

        tot_debito = float(df["debito"].sum())
        tot_credito = float(df["credito"].sum())
        saldo_final = saldo_anterior + tot_debito - tot_credito

        # Transformar en lista de diccionarios enriquecidos
        transactions = []
        for _, row in df.iterrows():
            transactions.append({
                "fecha": row["fecha"],
                "fecha_str": row["fecha"].strftime("%Y-%m-%d") if pd.notna(row["fecha"]) else "",
                "documento": row["documento"],
                "tipo_documento": str(row.get("tipo_documento", "")),
                "cuenta": int(row["cuenta"]),
                "descripcion_cuenta": str(row.get("descripcion_cuenta", cuenta_desc)),
                "detalle": str(row["detalle"]),
                "debito": float(row["debito"]),
                "credito": float(row["credito"]),
                # En contabilidad de bancos: débito = entrada (+), crédito = salida (-)
                "monto_neto": float(row["debito"] - row["credito"]),
                "tercero": str(row.get("tercero", "")) if pd.notna(row.get("tercero")) else "",
                "origen": "KARING",
                "archivo_origen": os.path.basename(file_path)
            })

        return {
            "archivo": file_path,
            "nombre_archivo": os.path.basename(file_path),
            "cuenta_codigo": cuenta_code,
            "cuenta_descripcion": cuenta_desc,
            "saldo_anterior": round(saldo_anterior, 2),
            "total_debito": round(tot_debito, 2),
            "total_credito": round(tot_credito, 2),
            "saldo_final": round(saldo_final, 2),
            "num_registros": len(transactions),
            "transacciones": transactions,
            "dataframe": df
        }
