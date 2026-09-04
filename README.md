# ✈️ Planetour S.A.S. - Sistema de Conciliación Bancaria vs Karing

Sistema automatizado de auditoría y conciliación contable desarrollado para **Planetour S.A.S.** Permite procesar mes a mes los extractos bancarios oficiales (PDF y reportes de pasarela) contra los libros auxiliares del sistema contable **Karing** (.xls/.xlsx).

---

## 🚀 Inicio Rápido

### Opción 1: Panel de Control Web Interactivo (Recomendado)
- Haz doble clic en el archivo **`iniciar_conciliador.bat`** en la carpeta principal.
- O ejecuta desde la terminal:
  ```bash
  python -m streamlit run app.py
  ```

### Opción 2: Ejecución por Línea de Comandos (CLI)
Para procesar meses por lotes o generar reportes automáticos:
```bash
# Conciliar un mes específico:
python reconcile_cli.py --month "AGOSTO 2026"

# Conciliar todos los meses y generar todos los Excel:
python reconcile_cli.py --all
```

---

## 🏦 Entidades Financieras y Cuentas Soportadas

1. **Bancolombia**:
   - Cuenta Operativa San Andrés (ADZ) #2173
   - Cuenta Operativa Leticia #3629
   - Cuenta Operativa Yopal #0220
2. **BBVA**:
   - Cuenta Corriente #3027
   - Cuenta Libretón #4073
   - Cuenta Ahorros #8357
3. **Banco de Bogotá**:
   - Cuenta Corriente #0884
4. **Davivienda**:
   - Cuenta Corriente #4051
   - Tarjetas de Crédito / Recaudos #1875
5. **Bold (Pasarela de Pagos)**:
   - Reportes mensuales de ventas y comisiones

---

## 🌟 Módulos y Funcionalidades Clave

- **📊 1. Tablero Ejecutivo de Control:** Resumen global de saldos, tasa de efectividad, diferencias brutas y cuadro por banco con filtros rápidos.
- **🚨 2. Lo que está MAL (Cruces Intercuentas):** Detección automática de dineros que ingresaron o salieron de un banco pero se registraron contablemente en otra cuenta en Karing, con instrucciones precisas de reclasificación.
- **❓ 3. Lo que FALTA por Conciliar:** Separación clara entre partidas no contabilizadas en Karing (pendientes de banco) y registros contables no reflejados en el extracto (pendientes de Karing).
- **💸 4. Gastos e Impuestos Bancarios:** Identificación automática de GMF (4x1000), comisiones ACH/Efecty/portales y rendimientos financieros con propuesta de asiento contable.
- **✅ 5. Conciliados 1 a 1 con ID Pareo:** Pareo exacto por valor, sentido y proximidad de fechas, asignando un código único (`PAR-XXXX-0001`) a cada cruce para trazabilidad total.
- **🔍 6. Buscador y Auditoría de Montos Repetidos:**
  - *Buscador Universal:* Localización por número de recibo, tercero, monto o palabra clave.
  - *Auditoría de Montos Repetidos:* Cuando un valor aparece decenas de veces (ej. pagos de $3.000.000), permite ver exactamente cuál movimiento cruzó y cuál quedó huérfano.
- **📤 7. Cargar Nuevos Periodos:** Carga interactiva para subir extractos y auxiliares de meses futuros.
- **📥 Descargas en Excel (.xlsx):** Todos los reportes, vistas individuales y el Libro Maestro Consolidado se exportan en formato Excel nativo con diseño corporativo.

---

## 📁 Estructura del Proyecto

```
conciliacion-planetour/
├── app.py                      # Interfaz web principal en Streamlit
├── reconcile_cli.py            # Ejecutor por consola / CLI
├── iniciar_conciliador.bat     # Lanzador de un clic para Windows
├── README.md                   # Documentación general del sistema
├── .gitignore                  # Reglas de exclusión de git
├── src/
│   ├── engine/
│   │   ├── account_map.py      # Catálogo y configuración de cuentas contables
│   │   └── matcher.py          # Motor de cruce, intercuentas y pareo 1 a 1
│   ├── parsers/
│   │   ├── bancolombia_parser.py
│   │   ├── bbva_parser.py
│   │   ├── bogota_parser.py
│   │   ├── bold_parser.py
│   │   ├── davivienda_parser.py
│   │   └── karing_parser.py
│   └── reports/
│       └── excel_generator.py  # Generador de libros Excel (.xlsx) formateados
├── BANCOLOMBIA/                # Extractos y auxiliares organizados por mes
├── BBVA/
├── BOGOTA/
├── BOLD/
├── DAVIVIENDA/
└── informes/                   # Salida de libros de conciliación generados
```

---

*Desarrollado para Planetour S.A.S.*
