"""
Parser para extractos bancarios de Banco de Bogotá en PDF.
"""

import pdfplumber
import re
from datetime import datetime
from typing import Dict, Any, List, Optional
import os

class BogotaParser:
    """
    Extrae transacciones, saldos, GMF y comisiones de extractos de Banco de Bogotá.
    """

    @staticmethod
    def parse_file(file_path: str) -> Dict[str, Any]:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"No se encontró el archivo: {file_path}")

        all_text = []
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text.append(text)

        full_content = "\n".join(all_text)
        lines = full_content.split("\n")

        cuenta_numero = ""
        periodo_texto = ""
        saldo_inicial = 0.0
        total_abonos = 0.0
        total_cargos = 0.0
        total_gmf = 0.0
        saldo_final = 0.0
        year = 2026

        for l in lines[:40]:
            l_clean = l.strip()
            # Cuenta
            m_cta = re.search(r'Cuenta\s*N[úu\?]mero:\s*(\d+)', l_clean, re.I)
            if m_cta:
                cuenta_numero = m_cta.group(1)

            # Periodo
            if "Desde:" in l_clean and "Hasta:" in l_clean:
                periodo_texto = l_clean

            # Resumen
            if "Saldo Inicial:" in l_clean:
                m = re.search(r'Saldo Inicial:\s*([-\d,]+\.\d{2})', l_clean)
                if m:
                    saldo_inicial = float(m.group(1).replace(",", ""))
            if "Total" in l_clean and "Abonos:" in l_clean:
                m = re.search(r'Abonos:\s*([-\d,]+\.\d{2})', l_clean)
                if m:
                    total_abonos = float(m.group(1).replace(",", ""))
            if "Total" in l_clean and "Cargos:" in l_clean:
                m = re.search(r'Cargos:\s*([-\d,]+\.\d{2})', l_clean)
                if m:
                    total_cargos = float(m.group(1).replace(",", ""))
            if "Total 4x1000 GMF:" in l_clean:
                m = re.search(r'Total 4x1000 GMF:\s*([-\d,]+\.\d{2})', l_clean)
                if m:
                    total_gmf = float(m.group(1).replace(",", ""))
            if "Saldo Final:" in l_clean:
                m = re.search(r'Saldo Final:\s*([-\d,]+\.\d{2})', l_clean)
                if m:
                    saldo_final = float(m.group(1).replace(",", ""))

        # Extraer transacciones
        # Formato: DD/MM [COD_TRANS] [DESCRIPCIÓN ...] [VALOR] [SALDO]
        tx_pattern = re.compile(r'^(\d{2}/\d{2})\s+([A-Z0-9]{4})\s+(.+?)\s+([-\d,]+\.\d{2})\s+([-\d,]+\.\d{2})$')

        transactions = []
        for l in lines:
            m = tx_pattern.match(l.strip())
            if m:
                fecha_part = m.group(1)
                cod_trans = m.group(2)
                desc = m.group(3).strip()
                val_num = float(m.group(4).replace(",", ""))
                saldo_num = float(m.group(5).replace(",", ""))

                d_parts = fecha_part.split("/")
                day = int(d_parts[0])
                month = int(d_parts[1])
                try:
                    fecha_dt = datetime(year, month, day)
                except Exception:
                    fecha_dt = None

                desc_upper = desc.upper()
                categoria = "OPERATIVA"
                if any(k in desc_upper for k in ["4X1.000", "GMF", "4X1000"]):
                    categoria = "GMF_4X1000"
                elif any(k in desc_upper for k in ["COMISION", "CUOTA DE MANEJO"]):
                    categoria = "COMISION"
                elif "INTERES" in desc_upper:
                    categoria = "INTERESES"
                elif "IVA" in desc_upper:
                    categoria = "IVA"

                transactions.append({
                    "fecha": fecha_dt,
                    "fecha_str": fecha_dt.strftime("%Y-%m-%d") if fecha_dt else f"{year}-{month:02d}-{day:02d}",
                    "documento": cod_trans,
                    "descripcion": desc,
                    "monto": val_num,
                    "es_abono": val_num > 0,
                    "es_cargo": val_num < 0,
                    "saldo": saldo_num,
                    "categoria": categoria,
                    "banco": "BOGOTA",
                    "cuenta_banco": cuenta_numero,
                    "archivo_origen": os.path.basename(file_path)
                })

        return {
            "banco": "BOGOTA",
            "archivo": file_path,
            "nombre_archivo": os.path.basename(file_path),
            "cuenta_numero": cuenta_numero,
            "periodo_texto": periodo_texto,
            "saldo_anterior": saldo_inicial,
            "total_abonos": total_abonos,
            "total_cargos": total_cargos,
            "saldo_final": saldo_final,
            "total_gmf": total_gmf,
            "num_transacciones": len(transactions),
            "transacciones": transactions
        }
