"""
Mapeo centralizado de cuentas bancarias y libros auxiliares de Karing para Planetour SAS.
"""

from dataclasses import dataclass
from typing import Optional, Dict, List

@dataclass
class AccountConfig:
    bank_name: str
    account_key: str
    account_number: str
    karing_code: int
    karing_name: str
    account_type: str  # Ahorros, Corriente, Pasarela
    folder_name: str
    subfolder_name: Optional[str] = None

# Definición del catálogo de cuentas de Planetour SAS
ACCOUNTS_CATALOG: Dict[str, AccountConfig] = {
    # BANCOLOMBIA
    "BANCOLOMBIA_ADZ": AccountConfig(
        bank_name="BANCOLOMBIA",
        account_key="BANCOLOMBIA_ADZ",
        account_number="34871212173",
        karing_code=11201002,
        karing_name="BANCOLOMBIA-SANANDRES",
        account_type="Ahorros",
        folder_name="BANCOLOMBIA",
        subfolder_name="CUENTA_ADZ_2173"
    ),
    "BANCOLOMBIA_LETICIA": AccountConfig(
        bank_name="BANCOLOMBIA",
        account_key="BANCOLOMBIA_LETICIA",
        account_number="94347513629",
        karing_code=11100504,
        karing_name="BANCOLOMBIA - PLANETOUR SAS",
        account_type="Corriente",
        folder_name="BANCOLOMBIA",
        subfolder_name="CUENTA_LETICIA_3629"
    ),
    "BANCOLOMBIA_YOPAL": AccountConfig(
        bank_name="BANCOLOMBIA",
        account_key="BANCOLOMBIA_YOPAL",
        account_number="36300040220",
        karing_code=11201005,
        karing_name="BANCOLOMBIA-YOPAL CUENTA DE AHORROS",
        account_type="Ahorros",
        folder_name="BANCOLOMBIA",
        subfolder_name="CUENTA_YOPAL_0220"
    ),
    
    # BBVA
    "BBVA_3027": AccountConfig(
        bank_name="BBVA",
        account_key="BBVA_3027",
        account_number="001305060200003027",
        karing_code=11201016,
        karing_name="BBVA CTA AHO GECKO 003027",
        account_type="Ahorros",
        folder_name="BBVA",
        subfolder_name="BBVA_3027"
    ),
    "BBVA_4073": AccountConfig(
        bank_name="BBVA",
        account_key="BBVA_4073",
        account_number="001305060200374073",
        karing_code=11201010,
        karing_name="BBVA CTA AHO  506-3740-73",
        account_type="Ahorros",
        folder_name="BBVA",
        subfolder_name="BBVA_4073"
    ),
    "BBVA_8357": AccountConfig(
        bank_name="BBVA",
        account_key="BBVA_8357",
        account_number="001305060100018357",
        karing_code=11100509,
        karing_name="BBVA CTA CTE 118357",
        account_type="Corriente",
        folder_name="BBVA",
        subfolder_name="BBVA_8357"
    ),
    
    # BANCO DE BOGOTA
    "BOGOTA_8097": AccountConfig(
        bank_name="BOGOTA",
        account_key="BOGOTA_8097",
        account_number="407228097",
        karing_code=11100501,
        karing_name="BANCO DE BOGOTA",
        account_type="Corriente",
        folder_name="BOGOTA",
        subfolder_name=None
    ),
    
    # DAVIVIENDA
    "DAVIVIENDA_4051": AccountConfig(
        bank_name="DAVIVIENDA",
        account_key="DAVIVIENDA_4051",
        account_number="286000214051",
        karing_code=11201008,
        karing_name="DAVIVIENDA CTA AHO 4051",
        account_type="Ahorros",
        folder_name="DAVIVIENDA",
        subfolder_name="CUENTA_4051"
    ),
    "DAVIVIENDA_1875": AccountConfig(
        bank_name="DAVIVIENDA",
        account_key="DAVIVIENDA_1875",
        account_number="1875",
        karing_code=11201004,
        karing_name="DAVIVIENDA TARJETAS 1875",
        account_type="Ahorros / Tarjetas",
        folder_name="DAVIVIENDA",
        subfolder_name="CUENTA_1875"
    ),
    
    # BOLD
    "BOLD_7997": AccountConfig(
        bank_name="BOLD",
        account_key="BOLD_7997",
        account_number="1700-1210-7997",
        karing_code=11201017,
        karing_name="BOLD CTA AHO 1700-1210-7997",
        account_type="Pasarela / Datáfonos",
        folder_name="BOLD",
        subfolder_name=None
    ),
}

# Índices rápidos por código Karing y número de cuenta
CODE_TO_ACCOUNT: Dict[int, AccountConfig] = {cfg.karing_code: cfg for cfg in ACCOUNTS_CATALOG.values()}

def get_account_by_karing_code(code: int) -> Optional[AccountConfig]:
    return CODE_TO_ACCOUNT.get(code)

def get_account_by_key(key: str) -> Optional[AccountConfig]:
    return ACCOUNTS_CATALOG.get(key)
