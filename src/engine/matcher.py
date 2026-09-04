"""
Motor de Conciliación Bancaria y Detección de Anomalías para Planetour SAS.
"""

from typing import Dict, Any, List, Optional, Tuple
import os
import glob
from datetime import datetime, timedelta
import pandas as pd

from src.engine.account_map import ACCOUNTS_CATALOG, AccountConfig, get_account_by_karing_code
from src.parsers.karing_parser import KaringParser
from src.parsers.bancolombia_parser import BancolombiaParser
from src.parsers.bbva_parser import BBVAParser
from src.parsers.bogota_parser import BogotaParser
from src.parsers.davivienda_parser import DaviviendaParser
from src.parsers.bold_parser import BoldParser

SPANISH_MONTH_TO_NUM = {
    "ENERO": "01", "FEBRERO": "02", "MARZO": "03", "ABRIL": "04",
    "MAYO": "05", "JUNIO": "06", "JULIO": "07", "AGOSTO": "08",
    "SEPTIEMBRE": "09", "OCTUBRE": "10", "NOVIEMBRE": "11", "DICIEMBRE": "12"
}

class ReconciliationEngine:
    """
    Motor central que orquesta la carga de archivos, cruces 1-a-1,
    detección de gastos bancarios/GMF, cruce intercuentas y auditoría de periodos.
    """

    def __init__(self, root_dir: str, date_tolerance_days: int = 4):
        self.root_dir = root_dir
        self.date_tolerance_days = date_tolerance_days

    def scan_available_months(self) -> List[str]:
        """Detecta los meses disponibles en la estructura de carpetas."""
        months = set()
        for folder in ["BANCOLOMBIA", "BBVA", "BOGOTA", "DAVIVIENDA", "BOLD"]:
            p = os.path.join(self.root_dir, folder)
            if os.path.exists(p):
                for sub in os.listdir(p):
                    sub_path = os.path.join(p, sub)
                    if os.path.isdir(sub_path) and any(m in sub.upper() for m in ["JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE", "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO"]):
                        months.add(sub.upper())
        return sorted(list(months))

    def reconcile_month(self, month_name: str) -> Dict[str, Any]:
        """
        Ejecuta la conciliación completa de todas las entidades para un mes determinado
        mediante un motor global en 3 fases:
        Fase 1: Segregación de gastos/impuestos bancarios y cruces directos 1-a-1 por cuenta.
        Fase 2: Cruce transversal intercuentas (registros en cuentas equivocadas). Marca tanto el banco
                como la partida de Karing de la otra cuenta para evitar doble contabilización.
        Fase 3: Generación de partidas pendientes (banco y karing) y estados formales sin solapamiento.
        """
        month_clean = month_name.strip().upper()
        warnings = []
        accounts_data = {}
        bank_txs_by_acc = {}
        karing_txs_by_acc = {}

        # 1. Carga de cada cuenta y sus archivos
        for acc_key, acc_cfg in ACCOUNTS_CATALOG.items():
            acc_result = self._load_account_files(acc_cfg, month_clean, warnings)
            accounts_data[acc_key] = acc_result

            bank = acc_result["bank"]
            karing = acc_result["karing"]

            b_txs = [dict(t) for t in bank["transacciones"]] if bank else []
            for i, bt in enumerate(b_txs):
                bt["id_interno"] = f"B_{acc_key}_{i}"
                bt["acc_key"] = acc_key
            bank_txs_by_acc[acc_key] = b_txs

            k_txs = [dict(t) for t in karing["transacciones"]] if karing else []
            for j, kt in enumerate(k_txs):
                kt["id_interno"] = f"K_{acc_key}_{j}"
                kt["acc_key"] = acc_key
            karing_txs_by_acc[acc_key] = k_txs

        # Estructuras de rastreo de uso global
        used_bank_by_acc = {k: set() for k in ACCOUNTS_CATALOG}
        used_karing_by_acc = {k: set() for k in ACCOUNTS_CATALOG}
        gastos_impuestos_by_acc = {k: [] for k in ACCOUNTS_CATALOG}
        conciliados_by_acc = {k: [] for k in ACCOUNTS_CATALOG}
        cruces_by_acc = {k: [] for k in ACCOUNTS_CATALOG}

        # -------------------------------------------------------------------------
        # FASE 1: Segregación de Gastos Bancarios y Cruce Directo 1-a-1 dentro de c/cta
        # -------------------------------------------------------------------------
        for acc_key, acc_cfg in ACCOUNTS_CATALOG.items():
            b_txs = bank_txs_by_acc[acc_key]
            k_txs = karing_txs_by_acc[acc_key]
            acc_res = accounts_data[acc_key]
            bank = acc_res["bank"]

            # 1.1 Gastos e impuestos bancarios (GMF, COMISION, INTERESES, IVA)
            for bt in b_txs:
                cat = bt.get("categoria", "OPERATIVA")
                if cat in ["GMF_4X1000", "COMISION", "INTERESES", "IVA"]:
                    used_bank_by_acc[acc_key].add(bt["id_interno"])
                    gastos_impuestos_by_acc[acc_key].append({
                        "tipo": cat,
                        "fecha": bt["fecha_str"],
                        "documento": bt.get("documento", ""),
                        "descripcion": bt["descripcion"],
                        "monto": bt["monto"],
                        "cuenta_banco": acc_cfg.karing_name,
                        "banco": acc_cfg.bank_name,
                        "sugerencia_contable": self._suggest_accounting_entry(cat, bt["monto"], acc_cfg)
                    })

            # 1.2 Comisiones y retenciones Bold consolidadas
            if acc_cfg.bank_name == "BOLD" and bank:
                if bank.get("total_comisiones", 0) > 0:
                    gastos_impuestos_by_acc[acc_key].append({
                        "tipo": "COMISION_BOLD",
                        "fecha": f"{month_clean_prefix(bank.get('archivo_origen', ''))}",
                        "documento": "CONSOLIDADO_BOLD",
                        "descripcion": "Comisiones y Deducciones Totales Bold",
                        "monto": -float(bank["total_comisiones"]),
                        "cuenta_banco": acc_cfg.karing_name,
                        "banco": "BOLD",
                        "sugerencia_contable": "Débito 530515 (Comisiones) vs Crédito 11201017 (Bold)"
                    })
                if bank.get("total_retefuente", 0) > 0:
                    gastos_impuestos_by_acc[acc_key].append({
                        "tipo": "RETEFUENTE_BOLD",
                        "fecha": f"{month_clean_prefix(bank.get('archivo_origen', ''))}",
                        "documento": "CONSOLIDADO_BOLD",
                        "descripcion": "Retención en la Fuente Practicada por Bold",
                        "monto": -float(bank["total_retefuente"]),
                        "cuenta_banco": acc_cfg.karing_name,
                        "banco": "BOLD",
                        "sugerencia_contable": "Débito 135515 (Anticipo ReteFuente) vs Crédito 11201017 (Bold)"
                    })

            # 1.3 Cruce directo 1-a-1 dentro de la misma cuenta
            for bt in b_txs:
                if bt["id_interno"] in used_bank_by_acc[acc_key]:
                    continue
                bt_val = round(abs(bt["monto"]), 2)
                bt_date = bt["fecha"]
                is_abono = bt["es_abono"]

                candidates = []
                for kt in k_txs:
                    if kt["id_interno"] in used_karing_by_acc[acc_key]:
                        continue
                    if is_abono and kt["debito"] > 0 and round(kt["debito"], 2) == bt_val:
                        pass
                    elif not is_abono and kt["credito"] > 0 and round(kt["credito"], 2) == bt_val:
                        pass
                    else:
                        continue

                    days_diff = 999
                    if bt_date and kt["fecha"] and pd.notna(kt["fecha"]):
                        days_diff = abs((bt_date - kt["fecha"]).days)
                    candidates.append((days_diff, kt))

                if candidates:
                    candidates.sort(key=lambda x: x[0])
                    best_diff, best_k = candidates[0]
                    if best_diff <= (self.date_tolerance_days + 3):
                        used_bank_by_acc[acc_key].add(bt["id_interno"])
                        used_karing_by_acc[acc_key].add(best_k["id_interno"])
                        conciliados_by_acc[acc_key].append({
                            "id_pareo": f"PAR-{acc_cfg.account_key}-{len(conciliados_by_acc[acc_key]) + 1:04d}",
                            "tipo_cruce": "DIRECTO_1_A_1",
                            "fecha_banco": bt["fecha_str"],
                            "documento_banco": bt.get("documento", ""),
                            "fecha_karing": best_k["fecha_str"],
                            "diferencia_dias": best_diff,
                            "documento_karing": best_k["documento"],
                            "tipo_documento_karing": best_k.get("tipo_documento", "RC"),
                            "detalle_karing": best_k["detalle"],
                            "descripcion_banco": bt["descripcion"],
                            "valor_conciliado": bt_val,
                            "sentido": "INGRESO" if is_abono else "EGRESO",
                            "cuenta_karing": acc_cfg.karing_name
                        })

        # -------------------------------------------------------------------------
        # FASE 2: Cruce Intercuentas (Reclasificaciones entre cuentas)
        # -------------------------------------------------------------------------
        for acc_key, acc_cfg in ACCOUNTS_CATALOG.items():
            b_txs = bank_txs_by_acc[acc_key]
            for bt in b_txs:
                if bt["id_interno"] in used_bank_by_acc[acc_key]:
                    continue
                if bt.get("categoria", "OPERATIVA") != "OPERATIVA":
                    continue

                bt_val = round(abs(bt["monto"]), 2)
                bt_date = bt["fecha"]
                is_abono = bt["es_abono"]

                candidates = []
                for other_acc_key, other_cfg in ACCOUNTS_CATALOG.items():
                    if other_acc_key == acc_key:
                        continue
                    for other_k in karing_txs_by_acc[other_acc_key]:
                        if other_k["id_interno"] in used_karing_by_acc[other_acc_key]:
                            continue
                        if is_abono and other_k["debito"] > 0 and round(other_k["debito"], 2) == bt_val:
                            pass
                        elif not is_abono and other_k["credito"] > 0 and round(other_k["credito"], 2) == bt_val:
                            pass
                        else:
                            continue

                        days_diff = 999
                        if bt_date and other_k["fecha"] and pd.notna(other_k["fecha"]):
                            days_diff = abs((bt_date - other_k["fecha"]).days)
                        if days_diff <= (self.date_tolerance_days + 2):
                            candidates.append((days_diff, other_acc_key, other_k))

                if candidates:
                    candidates.sort(key=lambda x: x[0])
                    best_diff, best_other_acc, best_k = candidates[0]
                    # Cruce confirmado: marcar AMBAS cuentas para que NO queden en pendientes
                    used_bank_by_acc[acc_key].add(bt["id_interno"])
                    used_karing_by_acc[best_other_acc].add(best_k["id_interno"])
                    cruces_by_acc[acc_key].append({
                        "tipo_hallazgo": "IMPUTACION_EQUIVOCADA_DE_CUENTA",
                        "fecha_banco": bt["fecha_str"],
                        "banco_real": acc_cfg.karing_name,
                        "descripcion_banco": bt["descripcion"],
                        "valor": bt_val,
                        "sentido": "INGRESO" if is_abono else "EGRESO",
                        "cuenta_registrada_karing": best_k["descripcion_cuenta"],
                        "codigo_cuenta_karing": best_k["cuenta"],
                        "documento_karing": best_k["documento"],
                        "fecha_karing": best_k["fecha_str"],
                        "detalle_karing": best_k["detalle"],
                        "accion_recomendada": f"Trasladar en Karing del auxiliar '{best_k['descripcion_cuenta']}' al auxiliar '{acc_cfg.karing_name}' (Doc #{best_k['documento']})"
                    })

        # -------------------------------------------------------------------------
        # FASE 3: Generación de Pendientes y Estados por Cuenta (Sin solapamientos)
        # -------------------------------------------------------------------------
        from collections import Counter
        results_by_account = {}

        for acc_key, acc_cfg in ACCOUNTS_CATALOG.items():
            acc_res = accounts_data[acc_key]
            bank = acc_res["bank"]
            karing = acc_res["karing"]

            if not bank and not karing:
                results_by_account[acc_key] = {
                    "config": acc_cfg,
                    "estado": "SIN_DATOS",
                    "metrics": {
                        "saldo_banco": 0.0,
                        "saldo_karing": 0.0,
                        "diferencia_bruta": 0.0,
                        "total_conciliados": 0,
                        "monto_conciliado": 0.0,
                        "total_gastos_bancarios": 0.0,
                        "total_gmf": 0.0,
                        "total_comisiones": 0.0,
                        "total_intereses": 0.0,
                        "num_cruces_intercuentas": 0,
                        "monto_cruces_intercuentas": 0.0,
                        "num_pendientes_banco": 0,
                        "monto_pendientes_banco": 0.0,
                        "num_pendientes_karing": 0,
                        "monto_pendientes_karing": 0.0
                    },
                    "conciliados": [],
                    "gastos_impuestos": [],
                    "cruces_intercuentas": [],
                    "pendientes_banco": [],
                    "pendientes_karing": [],
                    "bank_info": None,
                    "karing_info": None
                }
                continue

            b_txs = bank_txs_by_acc[acc_key]
            k_txs = karing_txs_by_acc[acc_key]
            acc_conciliados = conciliados_by_acc[acc_key]
            acc_gastos = gastos_impuestos_by_acc[acc_key]
            acc_cruces = cruces_by_acc[acc_key]

            banco_amounts_count = Counter([round(abs(b["monto"]), 2) for b in b_txs if b.get("categoria", "OPERATIVA") == "OPERATIVA"])
            karing_amounts_count = Counter([round(k["debito"] if k["debito"] > 0 else k["credito"], 2) for k in k_txs])
            matched_amounts_count = Counter([round(c["valor_conciliado"], 2) for c in acc_conciliados])

            # Pendientes Banco: no cruzados (ni directos, ni gastos, ni intercuentas)
            pendientes_banco = []
            for bt in b_txs:
                if bt["id_interno"] not in used_bank_by_acc[acc_key]:
                    val = round(abs(bt["monto"]), 2)
                    tot_b = banco_amounts_count.get(val, 1)
                    tot_k = karing_amounts_count.get(val, 0)
                    mat_c = matched_amounts_count.get(val, 0)
                    pend_b = tot_b - mat_c
                    if tot_b > 1 or tot_k > 1:
                        ctx_monto = f"Monto repetido: {tot_b} en Banco vs {tot_k} en Karing ({mat_c} cruzaron, {pend_b} pendiente en Banco)"
                    else:
                        ctx_monto = "Monto único en el mes"

                    pendientes_banco.append({
                        "fecha": bt["fecha_str"],
                        "documento": bt.get("documento", ""),
                        "descripcion": bt["descripcion"],
                        "monto": bt["monto"],
                        "tipo": "ABONO_NO_IDENTIFICADO" if bt["es_abono"] else "CARGO_NO_REGISTRADO",
                        "cuenta": acc_cfg.karing_name,
                        "contexto_monto": ctx_monto,
                        "total_mismo_monto_banco": tot_b,
                        "total_mismo_monto_karing": tot_k,
                        "conciliados_mismo_monto": mat_c
                    })

            # Pendientes Karing: no cruzados (ni directos ni como cuenta errónea intercuenta)
            pendientes_karing = []
            for kt in k_txs:
                if kt["id_interno"] not in used_karing_by_acc[acc_key]:
                    is_deb = kt["debito"] > 0
                    val = round(kt["debito"] if is_deb else kt["credito"], 2)
                    tot_b = banco_amounts_count.get(val, 0)
                    tot_k = karing_amounts_count.get(val, 1)
                    mat_c = matched_amounts_count.get(val, 0)
                    pend_k = tot_k - mat_c
                    if tot_b > 1 or tot_k > 1:
                        ctx_monto = f"Monto repetido: {tot_b} en Banco vs {tot_k} en Karing ({mat_c} cruzaron, {pend_k} pendiente en Karing)"
                    else:
                        ctx_monto = "Monto único en el mes"

                    pendientes_karing.append({
                        "fecha": kt["fecha_str"],
                        "documento": kt["documento"],
                        "tipo_documento": kt.get("tipo_documento", ""),
                        "detalle": kt["detalle"],
                        "debito": kt["debito"],
                        "credito": kt["credito"],
                        "tipo": "DEBITO_PENDIENTE_BANCO" if is_deb else "CREDITO_PENDIENTE_BANCO",
                        "cuenta": acc_cfg.karing_name,
                        "contexto_monto": ctx_monto,
                        "total_mismo_monto_banco": tot_b,
                        "total_mismo_monto_karing": tot_k,
                        "conciliados_mismo_monto": mat_c
                    })

            saldo_banco = float(bank["saldo_final"]) if bank else 0.0
            saldo_karing = float(karing["saldo_final"]) if karing else 0.0

            tot_gmf = sum(abs(g["monto"]) for g in acc_gastos if g["tipo"] == "GMF_4X1000")
            tot_comisiones = sum(abs(g["monto"]) for g in acc_gastos if "COMISION" in g["tipo"])
            tot_intereses = sum(abs(g["monto"]) for g in acc_gastos if g["tipo"] == "INTERESES")
            tot_gastos = sum(abs(g["monto"]) for g in acc_gastos if g["tipo"] != "INTERESES")
            diferencia_bruta = round(saldo_banco - saldo_karing, 2)

            results_by_account[acc_key] = {
                "config": acc_cfg,
                "estado": "CONCILIADA_CON_AJUSTES" if (pendientes_banco or pendientes_karing or acc_cruces or acc_gastos) else "CUADRADA",
                "metrics": {
                    "saldo_banco": saldo_banco,
                    "saldo_karing": saldo_karing,
                    "diferencia_bruta": diferencia_bruta,
                    "total_conciliados": len(acc_conciliados),
                    "monto_conciliado": round(sum(c["valor_conciliado"] for c in acc_conciliados), 2),
                    "total_gastos_bancarios": round(tot_gastos, 2),
                    "total_gmf": round(tot_gmf, 2),
                    "total_comisiones": round(tot_comisiones, 2),
                    "total_intereses": round(tot_intereses, 2),
                    "num_cruces_intercuentas": len(acc_cruces),
                    "monto_cruces_intercuentas": round(sum(c["valor"] for c in acc_cruces), 2),
                    "num_pendientes_banco": len(pendientes_banco),
                    "monto_pendientes_banco": round(sum(abs(p["monto"]) for p in pendientes_banco), 2),
                    "num_pendientes_karing": len(pendientes_karing),
                    "monto_pendientes_karing": round(sum(p["debito"] + p["credito"] for p in pendientes_karing), 2)
                },
                "conciliados": acc_conciliados,
                "gastos_impuestos": acc_gastos,
                "cruces_intercuentas": acc_cruces,
                "pendientes_banco": pendientes_banco,
                "pendientes_karing": pendientes_karing,
                "bank_info": bank,
                "karing_info": karing
            }

        # -------------------------------------------------------------------------
        # Consolidar métricas globales
        # -------------------------------------------------------------------------
        total_conciliados = sum(r["metrics"]["total_conciliados"] for r in results_by_account.values())
        monto_total_conciliado = sum(r["metrics"]["monto_conciliado"] for r in results_by_account.values())
        total_cruces_intercuentas = sum(len(r["cruces_intercuentas"]) for r in results_by_account.values())
        monto_total_cruces = sum(r["metrics"]["monto_cruces_intercuentas"] for r in results_by_account.values())
        total_gastos_bancarios = sum(r["metrics"]["total_gastos_bancarios"] for r in results_by_account.values())
        total_gmf = sum(r["metrics"]["total_gmf"] for r in results_by_account.values())
        total_comisiones = sum(r["metrics"]["total_comisiones"] for r in results_by_account.values())
        total_intereses = sum(r["metrics"]["total_intereses"] for r in results_by_account.values())
        total_pendientes_banco = sum(len(r["pendientes_banco"]) for r in results_by_account.values())
        monto_total_pendientes_banco = sum(r["metrics"]["monto_pendientes_banco"] for r in results_by_account.values())
        total_pendientes_karing = sum(len(r["pendientes_karing"]) for r in results_by_account.values())
        monto_total_pendientes_karing = sum(r["metrics"]["monto_pendientes_karing"] for r in results_by_account.values())

        saldo_total_bancos = sum(r["metrics"]["saldo_banco"] for r in results_by_account.values())
        saldo_total_karing = sum(r["metrics"]["saldo_karing"] for r in results_by_account.values())
        diferencia_global = saldo_total_bancos - saldo_total_karing

        total_resueltas = total_conciliados + total_cruces_intercuentas
        tasa_conciliacion = round((total_resueltas / (total_resueltas + total_pendientes_banco) * 100), 1) if (total_resueltas + total_pendientes_banco) > 0 else 100.0

        return {
            "mes": month_clean,
            "tolerancia_dias": self.date_tolerance_days,
            "advertencias": warnings,
            "cuentas": results_by_account,
            "resumen_global": {
                "num_cuentas": len(results_by_account),
                "total_conciliados": total_conciliados,
                "monto_total_conciliado": round(monto_total_conciliado, 2),
                "total_cruces_intercuentas": total_cruces_intercuentas,
                "monto_total_cruces": round(monto_total_cruces, 2),
                "total_gastos_bancarios": round(total_gastos_bancarios, 2),
                "total_gmf": round(total_gmf, 2),
                "total_comisiones": round(total_comisiones, 2),
                "total_intereses": round(total_intereses, 2),
                "total_pendientes_banco": total_pendientes_banco,
                "monto_total_pendientes_banco": round(monto_total_pendientes_banco, 2),
                "total_pendientes_karing": total_pendientes_karing,
                "monto_total_pendientes_karing": round(monto_total_pendientes_karing, 2),
                "saldo_total_bancos": round(saldo_total_bancos, 2),
                "saldo_total_karing": round(saldo_total_karing, 2),
                "diferencia_global": round(diferencia_global, 2),
                "tasa_conciliacion": tasa_conciliacion
            }
        }

    def _load_account_files(self, acc: AccountConfig, month: str, warnings: List[str]) -> Dict[str, Any]:
        """Localiza y procesa el extracto y el auxiliar Karing de una cuenta."""
        bank_res = None
        karing_res = None

        # Determinar ruta base del mes
        base_bank_dir = os.path.join(self.root_dir, acc.folder_name, month)
        if acc.subfolder_name:
            target_dir = os.path.join(base_bank_dir, acc.subfolder_name)
        else:
            target_dir = base_bank_dir

        # Para BOLD, los auxiliares también pueden estar directamente en la carpeta BOLD/mes o BOLD/
        if acc.bank_name == "BOLD":
            bold_aux = os.path.join(target_dir, f"LIBRO_AUXILIAR_{month.replace(' 2026', '')}_BOLD.xls")
            if not os.path.exists(bold_aux):
                bold_aux = os.path.join(self.root_dir, "BOLD", f"LIBRO_AUXILIAR_{month.replace(' 2026', '')}_BOLD.xls")
            if os.path.exists(bold_aux):
                try:
                    karing_res = KaringParser.parse_file(bold_aux)
                except Exception as e:
                    warnings.append(f"Error procesando Karing Bold {bold_aux}: {str(e)}")

        if os.path.exists(target_dir):
            for fname in os.listdir(target_dir):
                fpath = os.path.join(target_dir, fname)
                fname_l = fname.lower()
                
                # Archivo Karing (.xls)
                if fname_l.endswith(".xls") and "libro_auxiliar" in fname_l:
                    try:
                        karing_res = KaringParser.parse_file(fpath)
                    except Exception as e:
                        warnings.append(f"Error procesando Karing {fpath}: {str(e)}")

                # Archivo Extracto Banco (PDF o XLSX para Bold)
                if acc.bank_name == "BANCOLOMBIA" and fname_l.endswith(".pdf"):
                    try:
                        bank_res = BancolombiaParser.parse_file(fpath)
                        # Validar coherencia de periodo
                        if bank_res.get("periodo_hasta"):
                            p_hasta = bank_res["periodo_hasta"]
                            expected_num = next((num for m_txt, num in SPANISH_MONTH_TO_NUM.items() if m_txt in month), None)
                            if expected_num and f"/{expected_num}/" not in p_hasta:
                                warnings.append(
                                    f"Extracto Bancolombia '{fname}' en '{month}' (Cuenta: {acc.karing_name}) "
                                    f"corresponde internamente al período {p_hasta}. ¡El archivo está traspapelado o pertenece a otro mes!"
                                )
                    except Exception as e:
                        warnings.append(f"Error procesando Bancolombia {fpath}: {str(e)}")

                elif acc.bank_name == "BBVA" and fname_l.endswith(".pdf"):
                    try:
                        bank_res = BBVAParser.parse_file(fpath)
                    except Exception as e:
                        warnings.append(f"Error procesando BBVA {fpath}: {str(e)}")

                elif acc.bank_name == "BOGOTA" and fname_l.endswith(".pdf"):
                    try:
                        bank_res = BogotaParser.parse_file(fpath)
                    except Exception as e:
                        warnings.append(f"Error procesando Bogota {fpath}: {str(e)}")

                elif acc.bank_name == "DAVIVIENDA" and fname_l.endswith(".pdf"):
                    try:
                        bank_res = DaviviendaParser.parse_file(fpath)
                        # Validar coherencia de periodo
                        if bank_res.get("informe_mes"):
                            inf_mes = bank_res["informe_mes"].upper()
                            expected_month = next((m_txt for m_txt in SPANISH_MONTH_TO_NUM if m_txt in month), None)
                            if expected_month and expected_month not in inf_mes:
                                warnings.append(
                                    f"Extracto Davivienda '{fname}' en carpeta '{month}' (Cuenta: {acc.karing_name}) "
                                    f"corresponde internamente a '{inf_mes}'. ¡El archivo está traspapelado o pertenece a otro mes!"
                                )
                    except Exception as e:
                        warnings.append(f"Error procesando Davivienda {fpath}: {str(e)}")

                elif acc.bank_name == "BOLD" and fname_l.endswith(".xlsx") and "reporte" in fname_l:
                    try:
                        bank_res = BoldParser.parse_file(fpath)
                    except Exception as e:
                        warnings.append(f"Error procesando Bold {fpath}: {str(e)}")

        return {
            "account": acc,
            "bank": bank_res,
            "karing": karing_res
        }
    def _suggest_accounting_entry(self, tipo: str, monto: float, acc: AccountConfig) -> str:
        val = abs(monto)
        if tipo == "GMF_4X1000":
            return f"Débito 511595 (GMF 4x1000): ${val:,.2f} | Crédito {acc.karing_code} ({acc.karing_name}): ${val:,.2f}"
        elif tipo == "COMISION":
            return f"Débito 530515 (Comisiones Bancarias): ${val:,.2f} | Crédito {acc.karing_code} ({acc.karing_name}): ${val:,.2f}"
        elif tipo == "IVA":
            return f"Débito 240801 / 5305 (IVA en Gastos): ${val:,.2f} | Crédito {acc.karing_code} ({acc.karing_name}): ${val:,.2f}"
        elif tipo == "INTERESES":
            return f"Débito {acc.karing_code} ({acc.karing_name}): ${val:,.2f} | Crédito 421005 (Ingresos Intereses): ${val:,.2f}"
        return f"Ajustar en cuenta {acc.karing_code}: ${val:,.2f}"

def month_clean_prefix(filename: str) -> str:
    f = filename.lower()
    for m in ["agosto", "julio", "junio"]:
        if m in f:
            return f"2026-{m.upper()}"
    return "2026"
