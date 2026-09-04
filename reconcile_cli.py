"""
Interfaz de Línea de Comandos (CLI) para la Conciliación Bancaria de Planetour SAS.
Uso:
    python reconcile_cli.py --month "AGOSTO 2026"
    python reconcile_cli.py --all
"""

import argparse
import os
import sys
from src.engine.matcher import ReconciliationEngine
from src.reports.excel_generator import ExcelReportGenerator

def main():
    parser = argparse.ArgumentParser(description="Conciliador Bancario Automatizado - Planetour SAS")
    parser.add_argument("--month", type=str, default=None, help="Mes a conciliar (ej. 'AGOSTO 2026')")
    parser.add_argument("--all", action="store_true", help="Conciliar todos los meses disponibles")
    parser.add_argument("--tolerance", type=int, default=4, help="Días de tolerancia de fechas (def: 4)")
    parser.add_argument("--output_dir", type=str, default="informes", help="Directorio para guardar los informes")

    args = parser.parse_args()

    engine = ReconciliationEngine(root_dir=".", date_tolerance_days=args.tolerance)
    generator = ExcelReportGenerator()

    available_months = engine.scan_available_months()
    if not available_months:
        print("No se detectaron carpetas de meses en el directorio del proyecto.")
        sys.exit(1)

    if args.all:
        target_months = available_months
    elif args.month:
        target_months = [args.month.strip().upper()]
    else:
        # Por defecto el mes más reciente
        target_months = [available_months[-1]]

    os.makedirs(args.output_dir, exist_ok=True)

    print("=" * 70)
    print("SISTEMA DE CONCILIACIÓN BANCARIA AUTOMATIZADA - PLANETOUR S.A.S.")
    print(f"Meses a procesar: {', '.join(target_months)}")
    print(f"Tolerancia de fechas: ±{args.tolerance} días")
    print("=" * 70)

    for m in target_months:
        print(f"\nProcesando periodo: {m}...")
        res = engine.reconcile_month(m)

        # Imprimir alertas si existen
        if res["advertencias"]:
            print(f"  [!] Advertencias ({len(res['advertencias'])}):")
            for w in res["advertencias"]:
                print(f"      - {w}")

        glob = res["resumen_global"]
        print(f"  -> Conciliados exitosamente: {glob['total_conciliados']}")
        print(f"  -> Cruces intercuentas (cuenta errónea): {glob['total_cruces_intercuentas']}")
        print(f"  -> Gastos e Impuestos Bancarios: ${glob['total_gastos_bancarios']:,.2f} (GMF: ${glob['total_gmf']:,.2f}, Comisiones: ${glob['total_comisiones']:,.2f})")
        print(f"  -> Pendientes en Banco: {glob['total_pendientes_banco']} | Pendientes en Karing: {glob['total_pendientes_karing']}")

        out_filename = os.path.join(args.output_dir, f"Conciliacion_Planetour_{m.replace(' ', '_')}.xlsx")
        generator.generate_report(res, out_filename)
        print(f"  -> Informe Excel generado en: {out_filename}")

    print("\n" + "=" * 70)
    print("PROCESO DE CONCILIACIÓN FINALIZADO CON ÉXITO")
    print("=" * 70)

if __name__ == "__main__":
    main()
