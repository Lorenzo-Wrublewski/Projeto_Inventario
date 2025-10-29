# utils.py
# c:\Users\WRL1PO\Documents\Projeto_Inventario\@Parte 1\lib\utils.py
import os
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from .config import Config
from .logger import log
from .sap_actions import SapSession
from .waits import set_global_action_delay, wait_seconds, wait_for_interface_stable
from .storages import load_storages

def start_session(config: Config):
    pw = sync_playwright().start()
    browser = pw.chromium.launch(headless=config.headless, slow_mo=0)
    context = browser.new_context()
    page = context.new_page()
    page.set_default_navigation_timeout(config.nav_timeout_seconds * 1000)
    return pw, browser, context, page

def shutdown(pw, browser, context):
    for closeable in (context, browser):
        try:
            closeable.close()
        except Exception:
            pass
    try:
        pw.stop()
    except Exception:
        pass

def _esperar_sap_carregar(page, cfg: Config):
    tx_locator = lambda: page.get_by_role("textbox", name="Enter transaction code")
    tentativa = 0
    log("Iniciando carregamento do SAP (aguarda até campo pronto).")
    while True:
        tentativa += 1
        log(f"Tentativa {tentativa}: acessando {cfg.base_url}")
        try:
            page.goto(cfg.base_url, wait_until="domcontentloaded", timeout=cfg.nav_timeout_seconds * 1000)
        except TimeoutError:
            log(f"Timeout de navegação (>{cfg.nav_timeout_seconds}s). Vai repetir.", level="WARN")

        try:
            tx_locator().wait_for(timeout=cfg.wait_for_tx_field_seconds * 1000)
            log("Campo transação detectado.")
            wait_seconds(cfg.initial_stabilization_seconds, "Estabilização inicial")
            wait_for_interface_stable(
                page,
                timeout=cfg.interface_stable_timeout,
                min_stable=cfg.interface_stable_min_time
            )
            wait_seconds(cfg.wait_after_field_ready, "Pausa final antes da primeira transação")
            log("Interface pronta para uso.")
            return
        except Exception:
            log(f"Campo não apareceu em {cfg.wait_for_tx_field_seconds}s.", level="WARN")

        if cfg.max_retries_initial_load > 0 and tentativa >= cfg.max_retries_initial_load:
            log("Limite de tentativas atingido. Continuará em loop infinito até sucesso.", level="WARN")
            tentativa = 0

        wait_seconds(cfg.retry_delay_seconds, "Aguardando antes da nova tentativa")

def run_main():
    cfg = Config()
    set_global_action_delay(cfg.action_delay)

    # Garantir pasta de screenshots de erro
    os.makedirs(cfg.error_screenshot_dir, exist_ok=True)

    storages = load_storages(cfg.storages_csv_path)
    if not storages:
        log("Nenhum storage encontrado. Encerrando.", level="WARN")
        return

    pw, browser, context, page = start_session(cfg)
    try:
        _esperar_sap_carregar(page, cfg)
        session = SapSession(page, cfg)

        resultados = {}
        for st in storages:
            res = session.process_storage(st)
            resultados[st] = res

        log("Resumo execução storages:")
        for k, v in resultados.items():
            log(f"{k}: {v}")
    finally:
        shutdown(pw, browser, context)