#!/usr/bin/env bash
# ==============================================================================
# PLANETOUR S.A.S. - Actualizador de Cambios de Git y Reinicio de Servicio (Linux)
# ==============================================================================
set -e

GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\133[1;33m'
RED='\033[0;31m'
NC='\033[0m'

echo -e "\n${BLUE}======================================================${NC}"
echo -e "${BLUE}  PLANETOUR S.A.S. - Actualizando Servicio desde Git  ${NC}"
echo -e "${BLUE}======================================================${NC}"

CURRENT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$CURRENT_DIR"

# 1. Descargar últimos cambios del repositorio
echo -e "\n${YELLOW}[1/3] Descargando últimos cambios desde GitHub (git pull origin main)...${NC}"
git fetch origin
git pull origin main

# 2. Actualizar librerías de Python si hubo cambios en requirements.txt
echo -e "\n${YELLOW}[2/3] Verificando dependencias en venv...${NC}"
if [ -d "venv" ]; then
    ./venv/bin/pip install -r requirements.txt --quiet
else
    echo -e "${YELLOW}Creando entorno virtual venv...${NC}"
    python3 -m venv venv
    ./venv/bin/pip install --upgrade pip
    ./venv/bin/pip install -r requirements.txt
fi

# 3. Reiniciar servicio systemd
echo -e "\n${YELLOW}[3/3] Reiniciando servicio de conciliación...${NC}"
if systemctl is-active --quiet conciliacion.service 2>/dev/null; then
    sudo systemctl restart conciliacion.service
    echo -e "${GREEN}✓ Servicio conciliacion.service reiniciado correctamente.${NC}"
else
    echo -e "${YELLOW}Iniciando servicio de nuevo...${NC}"
    sudo systemctl restart conciliacion.service || sudo systemctl start conciliacion.service
fi

echo -e "\n${GREEN}======================================================${NC}"
echo -e "${GREEN}  ¡SISTEMA ACTUALIZADO Y EN FUNCIONAMIENTO!           ${NC}"
echo -e "${GREEN}======================================================${NC}"
sudo systemctl status conciliacion.service --no-pager -n 5
