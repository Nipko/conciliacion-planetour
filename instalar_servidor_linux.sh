#!/usr/bin/env bash
# ==============================================================================
# PLANETOUR S.A.S. - Instalador Automatizado en Servidor Linux (Ubuntu/Debian/RHEL)
# ==============================================================================
set -e

# Colores para salida visual
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # Sin color

echo -e "${BLUE}======================================================${NC}"
echo -e "${BLUE}  PLANETOUR S.A.S. - Instalación del Servicio Web     ${NC}"
echo -e "${BLUE}======================================================${NC}"

# 1. Detectar gestor de paquetes e instalar Git, Python3 y dependencias del sistema
echo -e "\n${YELLOW}[1/5] Verificando e instalando Git, Python3 y dependencias del sistema...${NC}"

if command -v apt-get &>/dev/null; then
    sudo apt-get update -y
    sudo apt-get install -y git python3 python3-pip python3-venv curl
elif command -v dnf &>/dev/null; then
    sudo dnf install -y git python3 python3-pip curl
elif command -v yum &>/dev/null; then
    sudo yum install -y git python3 python3-pip curl
else
    echo -e "${RED}No se detectó un gestor de paquetes soportado (apt, dnf o yum). Instale Git y Python3 manualmente.${NC}"
    exit 1
fi

echo -e "${GREEN}✓ Git y Python3 instalados correctamente.${NC}"

# 2. Determinar directorio del proyecto
CURRENT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$CURRENT_DIR"
echo -e "\n${YELLOW}[2/5] Directorio de instalación: $CURRENT_DIR${NC}"

# 3. Crear entorno virtual Python e instalar requerimientos
echo -e "\n${YELLOW}[3/5] Creando entorno virtual aislado (venv)...${NC}"
if [ ! -d "venv" ]; then
    python3 -m venv venv
fi

echo -e "${YELLOW}Instalando librerías requeridas (Streamlit, Pandas, OpenpyXL, PDFPlumber)...${NC}"
./venv/bin/pip install --upgrade pip
./venv/bin/pip install -r requirements.txt
echo -e "${GREEN}✓ Dependencias de Python instaladas con éxito.${NC}"

# 4. Configurar servicio de inicio automático (systemd)
echo -e "\n${YELLOW}[4/5] Configurando servicio systemd para ejecución 24/7 en segundo plano...${NC}"
SERVICE_FILE="/etc/systemd/system/conciliacion.service"
CURRENT_USER="$(whoami)"

sudo bash -c "cat <<EOF > $SERVICE_FILE
[Unit]
Description=Servicio de Conciliacion Bancaria Planetour
After=network.target

[Service]
Type=simple
User=$CURRENT_USER
WorkingDirectory=$CURRENT_DIR
ExecStart=$CURRENT_DIR/venv/bin/python -m streamlit run app.py --server.port=8501 --server.address=0.0.0.0 --server.headless=true
Restart=always
RestartSec=5
Environment=PYTHONUNBUFFERED=1

[Install]
WantedBy=multi-user.target
EOF"

echo -e "${GREEN}✓ Archivo de servicio creado en $SERVICE_FILE${NC}"

# 5. Habilitar e iniciar el servicio
echo -e "\n${YELLOW}[5/5] Iniciando y habilitando servicio conciliacion.service...${NC}"
sudo systemctl daemon-reload
sudo systemctl enable conciliacion.service
sudo systemctl restart conciliacion.service

# Obtener IP del servidor para mostrar al usuario
SERVER_IP=$(curl -s http://checkip.amazonaws.com || hostname -I | awk '{print $1}')

echo -e "\n${GREEN}======================================================${NC}"
echo -e "${GREEN}  ¡INSTALACIÓN COMPLETADA CON ÉXITO!                  ${NC}"
echo -e "${GREEN}======================================================${NC}"
echo -e "El servicio ya está activo y se ejecutará automáticamente al reiniciar el servidor."
echo -e ""
echo -e "🌐 Acceso Web al Sistema:"
echo -e "   - Desde la red local / navegador: ${BLUE}http://${SERVER_IP}:8501${NC}"
echo -e "   - En el servidor local:           ${BLUE}http://localhost:8501${NC}"
echo -e ""
echo -e "📋 Comandos útiles de administración:"
echo -e "   - Ver estado del servicio: ${YELLOW}sudo systemctl status conciliacion.service${NC}"
echo -e "   - Ver logs en vivo:        ${YELLOW}sudo journalctl -u conciliacion.service -f${NC}"
echo -e "   - Reiniciar servicio:      ${YELLOW}sudo systemctl restart conciliacion.service${NC}"
echo -e "   - Actualizar cambios Git:  ${YELLOW}./actualizar_servidor_linux.sh${NC}"
echo -e "======================================================\n"
