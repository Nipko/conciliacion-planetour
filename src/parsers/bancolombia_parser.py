"""
Parser para extractos bancarios de Bancolombia en PDF.
"""

import pdfplumber
import re
from datetime import datetime
from typing import Dict, Any, List, Optional
import os

class BancolombiaParser:
    """
    Extrae transacciones, saldos, resúmenes y metadatos de extractos Bancolombia.
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

        # Metadatos del encabezado
        cuenta_numero = ""
        periodo_desde = None
        periodo_hasta = None
        saldo_anterior = 0.0
        total_abonos = 0.0
        total_cargos = 0.0
        saldo_actual = 0.0
        intereses_pagados = 0.0

        for l in lines[:45]:
            l_clean = l.strip()
            # Número de cuenta
            m_cta = re.search(r'(?:N[ÚU\?]MERO|CUENTA.*?)\s*(\d{10,12})', l_clean, re.I)
            if m_cta and not cuenta_numero:
                cuenta_numero = m_cta.group(1)

            # Periodo DESDE: YYYY/MM/DD HASTA: YYYY/MM/DD
            m_per = re.search(r'DESDE:\s*(\d{4}[/-]\d{2}[/-]\d{2})\s+HASTA:\s*(\d{4}[/-]\d{2}[/-]\d{2})', l_clean, re.I)
            if m_per:
                periodo_desde = m_per.group(1).replace("-", "/")
                periodo_hasta = m_per.group(2).replace("-", "/")

            # Resumen de Saldos
            if "SALDO ANTERIOR" in l_clean:
                m_sa = re.search(r'SALDO ANTERIOR\s*\$\s*([-\d,]+\.\d{2})', l_clean)
                if m_sa:
                    saldo_anterior = float(m_sa.group(1).replace(",", ""))
            if "TOTAL ABONOS" in l_clean:
                m_ab = re.search(r'TOTAL ABONOS\s*\$\s*([-\d,]+\.\d{2})', l_clean)
                if m_ab:
                    total_abonos = float(m_ab.group(1).replace(",", ""))
            if "TOTAL CARGOS" in l_clean:
                m_cg = re.search(r'TOTAL CARGOS\s*\$\s*([-\d,]+\.\d{2})', l_clean)
                if m_cg:
                    total_cargos = float(m_cg.group(1).replace(",", ""))
            if "SALDO ACTUAL" in l_clean:
                m_sac = re.search(r'SALDO ACTUAL\s*\$\s*([-\d,]+\.\d{2})', l_clean)
                if m_sac:
                    saldo_actual = float(m_sac.group(1).replace(",", ""))
            if "VALOR INTERESES PAGADOS" in l_clean:
                m_int = re.search(r'VALOR INTERESES PAGADOS\s*\$\s*([-\d,]+\.\d{2})', l_clean)
                if m_int:
                    intereses_pagados = float(m_int.group(1).replace(",", ""))

        # Determinar año base a partir de periodo_hasta
        year = 2026
        if periodo_hasta:
            try:
                year = int(periodo_hasta.split("/")[0])
            except Exception:
                pass

        # Extracción de transacciones
        transactions = []
        # Patrón: D/MM o DD/MM [DESCRIPCIÓN ...] [VALOR] [SALDO]
        tx_pattern = re.compile(r'^(\d{1,2}/\d{2})\s+(.+?)\s+([-\d,]+\.\d{2})\s+([-\d,]+\.\d{2})$')

        for l in lines:
            l_str = l.strip()
            m = tx_pattern.match(l_str)
            if m:
                fecha_part = m.group(1)
                desc = m.group(2).strip()
                val_num = float(m.group(3).replace(",", ""))
                saldo_num = float(m.group(4).replace(",", ""))

                # Formatear fecha
                d_parts = fecha_part.split("/")
                day = int(d_parts[0])
                month = int(d_parts[1])
                try:
                    fecha_dt = datetime(year, month, day)
                except Exception:
                    fecha_dt = None

                # Categorizar conceptos bancarios
                desc_upper = desc.upper()
                categoria = "OPERATIVA"
                if "INTERES" in desc_upper:
                    categoria = "INTERESES"
                elif any(k in desc_upper for k in ["4X1000", "GMF", "GRAVAMEN"]):
                    categoria = "GMF_4X1000"
                elif any(k in desc_upper for k in ["COMISION", "CUOTA DE MANEJO"]):
                    categoria = "COMISION"
                elif "IVA" in desc_upper:
                    categoria = "IVA"

                transactions.append({
                    "fecha": fecha_dt,
                    "fecha_str": fecha_dt.strftime("%Y-%m-%d") if fecha_dt else f"{year}-{month:02d}-{day:02d}",
                    "descripcion": desc,
                    "monto": val_num,
                    "es_abono": val_num > 0,
                    "es_cargo": val_num < 0,
                    "saldo": saldo_num,
                    "categoria": categoria,
                    "banco": "BANCOLOMBIA",
                    "cuenta_banco": cuenta_numero,
                    "archivo_origen": os.path.basename(file_path)
                })

        return {
            "banco": "BANCOLOMBIA",
            "archivo": file_path,
            "nombre_archivo": os.path.basename(file_path),
            "cuenta_numero": cuenta_numero,
            "periodo_desde": periodo_desde,
            "periodo_hasta": periodo_hasta,
            "saldo_anterior": saldo_anterior,
            "total_abonos": total_abonos,
            "total_cargos": total_cargos,
            "saldo_final": saldo_actual,
            "intereses_pagados": intereses_pagados,
            "num_transacciones": len(transactions),
            "transacciones": transactions
        }
