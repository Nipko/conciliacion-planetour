"""
Parser para extractos bancarios de Davivienda en PDF.
"""

import pdfplumber
import re
from datetime import datetime
from typing import Dict, Any, List, Optional
import os

class DaviviendaParser:
    """
    Extrae transacciones, GMF, comisiones e IVA de extractos Davivienda.
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
        informe_mes = ""
        saldo_anterior = 0.0
        mas_creditos = 0.0
        menos_debitos = 0.0
        nuevo_saldo = 0.0
        year = 2026

        for l in lines[:35]:
            l_clean = l.strip()
            l_nodigits = re.sub(r'\s+', '', l_clean)
            
            # Cuenta (admite con y sin espacios, ej: 2860 0021 4051 o 286000214051)
            if re.match(r'^\d{10,14}$', l_nodigits) and not cuenta_numero:
                cuenta_numero = l_nodigits

            # INFORME DEL MES: MES / YYYY (manejando tipografía desfasada como 'ME SJ:UNIO' o 'INFORMEDELMES:')
            l_norm_header = re.sub(r'[\s:;]+', '', l_clean.upper())
            if ("INFORMEDELMES" in l_norm_header or "INFORME" in l_clean.upper()) and not informe_mes:
                for m in ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]:
                    if m in l_norm_header:
                        informe_mes = m
                        break
                m_yr = re.search(r'20\d{2}', l_norm_header)
                if m_yr:
                    year = int(m_yr.group(0))

            # Resumen saldos
            if "Saldo" in l_clean and "Anterior" in l_clean:
                m = re.search(r'Saldo\s*Anterior\s*\$\s*([-\d,]+\.\d{2})', l_clean, re.IGNORECASE)
                if m:
                    saldo_anterior = float(m.group(1).replace(",", ""))
            if "Cr" in l_clean and "ditos" in l_clean:  # MásCréditos
                m = re.search(r'ditos\s*\$\s*([-\d,]+\.\d{2})', l_clean, re.IGNORECASE)
                if m:
                    mas_creditos = float(m.group(1).replace(",", ""))
            if "bitos" in l_clean:  # MenosDébitos
                m = re.search(r'bitos\s*\$\s*([-\d,]+\.\d{2})', l_clean, re.IGNORECASE)
                if m:
                    menos_debitos = float(m.group(1).replace(",", ""))
            if "Nuevo" in l_clean and "Saldo" in l_clean:
                m = re.search(r'Nuevo\s*Saldo\s*\$\s*([-\d,]+\.\d{2})', l_clean, re.IGNORECASE)
                if m:
                    nuevo_saldo = float(m.group(1).replace(",", ""))

        # Si cuenta_numero sigue vacía, deducir por la ruta o carpeta
        if not cuenta_numero:
            if "4051" in file_path:
                cuenta_numero = "286000214051"
            elif "1875" in file_path:
                cuenta_numero = "266000311875"

        # Extraer transacciones
        # Formato 1 (Ahorros / Portafolio - 4051): DD MM $ VALOR[+-] [DOC] [DESCRIPCION]
        tx_pattern_fmt1 = re.compile(r'^(\d{2})\s+(\d{2})\s+\$\s*([-\d,]+\.\d{2})([+-])\s+(\d+)\s+(.*)$')
        
        # Formato 2 (Damas / Tarjetas - 1875): DD MM OFICINA DESCRIPCION DOC $ DEBITO $ CREDITO
        tx_pattern_fmt2 = re.compile(r'^(\d{2})\s+(\d{2})\s+(\d+)\s+(.+?)\s+(\d+)\s+\$\s*([\d,]+\.\d{2})\s+\$\s*([\d,]+\.\d{2})')

        transactions = []
        for l in lines:
            line_str = l.strip()
            
            # Probar Formato 1
            m1 = tx_pattern_fmt1.match(line_str)
            if m1:
                day = int(m1.group(1))
                month = int(m1.group(2))
                val_abs = float(m1.group(3).replace(",", ""))
                sign = m1.group(4)
                doc = m1.group(5)
                desc = m1.group(6).strip()

                try:
                    fecha_dt = datetime(year, month, day)
                except Exception:
                    fecha_dt = None

                signed_monto = val_abs if sign == "+" else -val_abs
                desc_upper = desc.upper()

                categoria = "OPERATIVA"
                if any(k in desc_upper for k in ["GRAVAMEN", "4X1000", "GMF"]):
                    categoria = "GMF_4X1000"
                elif bool(re.search(r'\bIVA\b', desc_upper)) and sign == "-":
                    categoria = "IVA"
                elif any(k in desc_upper for k in ["COBRO", "COMISION", "CUOTA"]):
                    categoria = "COMISION"
                elif "INTERES" in desc_upper:
                    categoria = "INTERESES"

                transactions.append({
                    "fecha": fecha_dt,
                    "fecha_str": fecha_dt.strftime("%Y-%m-%d") if fecha_dt else f"{year}-{month:02d}-{day:02d}",
                    "documento": doc,
                    "descripcion": desc,
                    "monto": signed_monto,
                    "es_abono": sign == "+",
                    "es_cargo": sign == "-",
                    "saldo": None,
                    "categoria": categoria,
                    "banco": "DAVIVIENDA",
                    "cuenta_banco": cuenta_numero,
                    "archivo_origen": os.path.basename(file_path)
                })
                continue

            # Probar Formato 2
            m2 = tx_pattern_fmt2.match(line_str)
            if m2:
                day = int(m2.group(1))
                month = int(m2.group(2))
                desc = m2.group(4).strip()
                doc = m2.group(5)
                debito = float(m2.group(6).replace(",", ""))
                credito = float(m2.group(7).replace(",", ""))

                try:
                    fecha_dt = datetime(year, month, day)
                except Exception:
                    fecha_dt = None

                if credito > 0:
                    signed_monto = credito
                    sign = "+"
                else:
                    signed_monto = -debito
                    sign = "-"

                desc_upper = desc.upper()

                categoria = "OPERATIVA"
                if any(k in desc_upper for k in ["GRAVAMEN", "4X1000", "GMF"]):
                    categoria = "GMF_4X1000"
                elif bool(re.search(r'\bIVA\b', desc_upper)) and sign == "-":
                    categoria = "IVA"
                elif any(k in desc_upper for k in ["COBRO", "COMISION", "CUOTA"]):
                    categoria = "COMISION"
                elif "RENDIMIENTO" in desc_upper or "INTERES" in desc_upper:
                    categoria = "INTERESES"

                transactions.append({
                    "fecha": fecha_dt,
                    "fecha_str": fecha_dt.strftime("%Y-%m-%d") if fecha_dt else f"{year}-{month:02d}-{day:02d}",
                    "documento": doc,
                    "descripcion": desc,
                    "monto": signed_monto,
                    "es_abono": sign == "+",
                    "es_cargo": sign == "-",
                    "saldo": None,
                    "categoria": categoria,
                    "banco": "DAVIVIENDA",
                    "cuenta_banco": cuenta_numero,
                    "archivo_origen": os.path.basename(file_path)
                })

        return {
            "banco": "DAVIVIENDA",
            "archivo": file_path,
            "nombre_archivo": os.path.basename(file_path),
            "cuenta_numero": cuenta_numero,
            "informe_mes": informe_mes,
            "saldo_anterior": saldo_anterior,
            "total_abonos": mas_creditos,
            "total_cargos": menos_debitos,
            "saldo_final": nuevo_saldo,
            "num_transacciones": len(transactions),
            "transacciones": transactions
        }
