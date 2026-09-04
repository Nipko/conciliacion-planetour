"""
Parser para extractos bancarios de BBVA en PDF.
"""

import pdfplumber
import re
from datetime import datetime
from typing import Dict, Any, List, Optional
import os

class BBVAParser:
    """
    Extrae transacciones, 4x1000, comisiones y saldos de extractos BBVA.
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
        periodo_desde = None
        periodo_hasta = None
        saldo_anterior = 0.0
        total_abonos = 0.0
        total_cargos = 0.0
        saldo_final = 0.0
        total_gmf = 0.0
        total_intereses = 0.0

        for l in lines[:40]:
            l_clean = l.strip()
            # Cuenta (001305...)
            m_cta = re.search(r'\b(0013\d{14,16})\b', l_clean)
            if m_cta and not cuenta_numero:
                cuenta_numero = m_cta.group(1)

            # Periodo DESDE: DD-MM-YYYY HASTA: DD-MM-YYYY
            m_per = re.search(r'PER[ÍI\?]ODO DESDE:\s*(\d{2}-\d{2}-\d{4})\s+HASTA:\s*(\d{2}-\d{2}-\d{4})', l_clean, re.I)
            if m_per:
                periodo_desde = m_per.group(1)
                periodo_hasta = m_per.group(2)

            # Resumen
            if "SALDO CIERRE MES ANTERIOR" in l_clean:
                m = re.search(r'SALDO CIERRE MES ANTERIOR\s*([-\d,]+\.\d{2})', l_clean)
                if m:
                    saldo_anterior = float(m.group(1).replace(",", ""))
            if "+ ABONOS" in l_clean:
                m = re.search(r'\+\s*ABONOS\s+\d+\s+([-\d,]+\.\d{2})', l_clean)
                if m:
                    total_abonos = float(m.group(1).replace(",", ""))
            if "- CARGOS" in l_clean:
                m = re.search(r'-\s*CARGOS\s+\d+\s+([-\d,]+\.\d{2})', l_clean)
                if m:
                    total_cargos = float(m.group(1).replace(",", ""))
            if "- 4 POR MIL" in l_clean:
                m = re.search(r'-\s*4 POR MIL\s+\d*\s*([-\d,]+\.\d{2})', l_clean)
                if m:
                    total_gmf = float(m.group(1).replace(",", ""))
            if "+ INTERESES RECIBIDOS" in l_clean:
                m = re.search(r'\+\s*INTERESES RECIBIDOS\s+\d*\s*([-\d,]+\.\d{2})', l_clean)
                if m:
                    total_intereses = float(m.group(1).replace(",", ""))
            if "SALDO FINAL" in l_clean:
                m = re.search(r'SALDO FINAL\s*([-\d,]+\.\d{2})', l_clean)
                if m:
                    saldo_final = float(m.group(1).replace(",", ""))

        # Extraer transacciones
        # Patrón: [DOC] DD-MM-YYYY DD-MM-YYYY [DESCRIPCIÓN] [VALOR] [SALDO]
        tx_pattern = re.compile(r'^(\d+)\s+(\d{2}-\d{2}-\d{4})\s+(\d{2}-\d{2}-\d{4})\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$')

        raw_txs = []
        for l in lines:
            m = tx_pattern.match(l.strip())
            if m:
                raw_txs.append({
                    "doc": m.group(1),
                    "fecha_op": m.group(2),
                    "fecha_val": m.group(3),
                    "desc": m.group(4).strip(),
                    "valor": float(m.group(5).replace(",", "")),
                    "saldo": float(m.group(6).replace(",", ""))
                })

        # Determinar signo con saldo sucesivo y descripción
        transactions = []
        prev_saldo = saldo_anterior

        for i, tx in enumerate(raw_txs):
            val_abs = tx["valor"]
            current_saldo = tx["saldo"]
            desc_upper = tx["desc"].upper()

            # Clasificación de cargo vs abono mediante cálculo matemático exacto del saldo
            if i > 0:
                diff = round(current_saldo - raw_txs[i-1]["saldo"], 2)
            else:
                diff = round(current_saldo - saldo_anterior, 2)

            if diff < 0:
                is_cargo = True
            elif diff > 0:
                is_cargo = False
            else:
                # Fallback solo si diff == 0 (caso extremo)
                if desc_upper.startswith("ABONO") or any(k in desc_upper for k in ["RECIBISTE", "INTERESES", "CONSIGNACION"]):
                    is_cargo = False
                else:
                    is_cargo = True

            signed_monto = -val_abs if is_cargo else val_abs

            # Parse fecha (usamos fecha_val como la fecha contable efectiva)
            try:
                fecha_dt = datetime.strptime(tx["fecha_val"], "%d-%m-%Y")
            except Exception:
                try:
                    fecha_dt = datetime.strptime(tx["fecha_op"], "%d-%m-%Y")
                except Exception:
                    fecha_dt = None

            # Categorización
            categoria = "OPERATIVA"
            if any(k in desc_upper for k in ["4X1.000", "4 POR MIL", "GMF"]):
                categoria = "GMF_4X1000"
            elif any(k in desc_upper for k in ["COMISION", "CUOTA DE MANEJO"]):
                categoria = "COMISION"
            elif "INTERES" in desc_upper:
                categoria = "INTERESES"
            elif "IVA" in desc_upper:
                categoria = "IVA"

            transactions.append({
                "fecha": fecha_dt,
                "fecha_str": fecha_dt.strftime("%Y-%m-%d") if fecha_dt else tx["fecha_val"],
                "documento": tx["doc"],
                "descripcion": tx["desc"],
                "monto": signed_monto,
                "es_abono": not is_cargo,
                "es_cargo": is_cargo,
                "saldo": current_saldo,
                "categoria": categoria,
                "banco": "BBVA",
                "cuenta_banco": cuenta_numero,
                "archivo_origen": os.path.basename(file_path)
            })

        return {
            "banco": "BBVA",
            "archivo": file_path,
            "nombre_archivo": os.path.basename(file_path),
            "cuenta_numero": cuenta_numero,
            "periodo_desde": periodo_desde,
            "periodo_hasta": periodo_hasta,
            "saldo_anterior": saldo_anterior,
            "total_abonos": total_abonos,
            "total_cargos": total_cargos,
            "saldo_final": saldo_final,
            "total_gmf": total_gmf,
            "total_intereses": total_intereses,
            "num_transacciones": len(transactions),
            "transacciones": transactions
        }
