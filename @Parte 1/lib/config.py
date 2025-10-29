# config.py
# c:\Users\WRL1PO\Documents\Projeto_Inventario\@Parte 1\lib\config.py
from dataclasses import dataclass

@dataclass
class Config:
    base_url: str = "http://rb3qraa0.server.bosch.com:8001/sap/bc/gui/sap/its/webgui/#"
    variant_name: str = "MMS3CA"
    storage_code: str = "J0A"  # Compatibilidade
    wait_default: float = 3.0
    wait_short: float = 1.0
    wait_long: float = 10.0
    wait_after_f8_seconds: int = 60
    screenshot_after_f8: str = "screenshot_storage.png"
    headless: bool = False
    action_delay: float = 2.0

    # Carregamento inicial
    nav_timeout_seconds: int = 120
    wait_for_tx_field_seconds: int = 300
    retry_delay_seconds: int = 15
    max_retries_initial_load: int = 0

    # Estabilização
    initial_stabilization_seconds: float = 3.0
    interface_stable_timeout: float = 90.0
    interface_stable_min_time: float = 2.0
    wait_tx_ready_timeout: float = 120.0
    wait_after_field_ready: float = 1.5

    # Loop storages
    storages_csv_path: str = "storages.csv"
    error_screenshot_dir: str = "error_screenshots"

    # Robustez LX15 / F8
    f8_retry_attempts: int = 5
    f8_retry_interval: float = 4.0
    reenter_storage_on_retry: bool = True
    quick_detection_timeout: float = 5.0      # Tempo máximo (s) da checagem rápida por tentativa
    quick_detection_poll: float = 0.5         # Intervalo de polling na detecção rápida
    wait_post_f8_small: float = 1.2           # Pausa curta logo após F8 antes da checagem rápida
    final_full_detection: bool = True         # Última tentativa usa detecção completa (timeout longo)
    f8_focus_retry: bool = True               # Reforça F8 via teclado global se necessário