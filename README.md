# ✈️ Planetour S.A.S. - Sistema de Conciliación Bancaria vs Karing

Sistema automatizado de auditoría y conciliación contable desarrollado para **Planetour S.A.S.** Permite procesar mes a mes los extractos bancarios oficiales (PDF y reportes de pasarela) contra los libros auxiliares del sistema contable **Karing** (.xls/.xlsx).

---

## 🖥️ Instalación en un Nuevo Servidor / Equipo (Windows 11)

En un nuevo equipo con Windows 11, solo debes seguir estos 3 pasos:

### Paso 1: Clonar el repositorio
Abre **PowerShell** o **Símbolo del sistema (CMD)** y ejecuta:
```bash
git clone https://github.com/Nipko/conciliacion-planetour.git
cd conciliacion-planetour
```
*(Si el nuevo equipo aún no tiene Git, puedes descargar el ZIP desde GitHub o ejecutar `winget install --id Git.Git -e --source winget` en PowerShell).*

### Paso 2: Ejecutar el Instalador Automático
Haz doble clic en el archivo:
👉 **`instalar_windows11.bat`**

Este script hace todo automáticamente:
1. Verifica si **Git** está instalado (y si no, lo instala de inmediato con `winget`).
2. Verifica si **Python** está instalado (y si no, lo instala automáticamente).
3. Crea el entorno virtual aislado (`venv`).
4. Instala todas las librerías necesarias (`requirements.txt`).
5. Habilita el puerto 8501 en el Firewall de Windows para que otros computadores de la oficina puedan entrar.
6. Inicia el panel web inmediatamente.

---

## 🔄 Cómo Actualizar Cambios desde GitHub (Windows 11)

Cada vez que se suban mejoras o cambios al repositorio en GitHub, ve a la carpeta del proyecto y haz doble clic en:
👉 **`actualizar_windows11.bat`**

El script automáticamente:
1. Descarga el código más reciente (`git pull origin main`).
2. Actualiza cualquier nueva librería de Python.
3. Inicia el sistema con las nuevas funciones aplicadas.

---

## 🚀 Uso Diario en Windows
Para abrir la aplicación en el día a día, simplemente haz doble clic en:
👉 **`iniciar_conciliador.bat`**

- Abrirá el navegador web en `http://localhost:8501`.
- Los demás equipos de la red pueden ingresar mediante: `http://[IP-DEL-EQUIPO]:8501`.

---

## 🐧 Instalación en Servidores Linux (Ubuntu / Debian / RHEL)
Si deseas desplegarlo en un servidor Linux o en la nube (AWS, Azure, DigitalOcean):
```bash
# 1. Instalar y configurar como servicio 24/7 (systemd):
chmod +x instalar_servidor_linux.sh
./instalar_servidor_linux.sh

# 2. Para actualizar cambios futuros en Linux:
chmod +x actualizar_servidor_linux.sh
./actualizar_servidor_linux.sh
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
- **🚨 2. Lo que está MAL (Cruces Intercuentas):** Detección automática de dineros que entraron/salieron de un banco pero se registraron contablemente en otra cuenta en Karing, con instrucciones precisas de reclasificación.
- **❓ 3. Lo que FALTA por Conciliar:** Separación clara entre partidas no contabilizadas en Karing (pendientes de banco) y registros contables no reflejados en el extracto (pendientes de Karing).
- **💸 4. Gastos e Impuestos Bancarios:** Identificación automática de GMF (4x1000), comisiones ACH/Efecty/portales y rendimientos financieros con propuesta de asiento contable.
- **✅ 5. Conciliados 1 a 1 con ID Pareo:** Pareo exacto por valor, sentido y proximidad de fechas, asignando un código único (`PAR-XXXX-0001`) a cada cruce para trazabilidad total.
- **🔍 6. Buscador y Auditoría de Montos Repetidos:**
  - *Buscador Universal:* Localización por número de recibo, tercero, monto o palabra clave.
  - *Auditoría de Montos Repetidos:* Cuando un valor aparece decenas de veces (ej. pagos de $3.000.000), permite ver exactamente cuál movimiento cruzó y cuál quedó huérfano.
- **📤 7. Cargar Nuevos Periodos:** Carga interactiva para subir extractos y auxiliares de meses futuros.
- **📥 Descargas en Excel (.xlsx):** Todos los reportes, vistas individuales y el Libro Maestro Consolidado se exportan en formato Excel nativo con diseño corporativo.

---

## 📁 Estructura del Repositorio

```
conciliacion-planetour/
├── app.py                         # Interfaz web principal en Streamlit
├── reconcile_cli.py               # Ejecutor por consola / CLI
├── requirements.txt               # Dependencias Python
├── iniciar_conciliador.bat        # Lanzador diario para Windows
├── instalar_windows11.bat         # Instalador automático para Windows 11
├── actualizar_windows11.bat       # Actualizador Git para Windows 11
├── instalar_servidor_linux.sh     # Instalador automático con systemd (Linux)
├── actualizar_servidor_linux.sh   # Actualizador Git para Linux
├── README.md                      # Documentación general del sistema
├── .gitignore                     # Reglas de exclusión de git
├── src/
│   ├── engine/                    # Motores de cruce, mapeo y balance
│   ├── parsers/                   # Extractores de extractos PDF y Karing
│   └── reports/                   # Generador de libros Excel (.xlsx)
├── BANCOLOMBIA/                   # Extractos y libros auxiliares por mes
├── BBVA/
├── BOGOTA/
├── BOLD/
├── DAVIVIENDA/
└── informes/                      # Libros de conciliación generados
```

---

*Desarrollado para Planetour S.A.S.*
