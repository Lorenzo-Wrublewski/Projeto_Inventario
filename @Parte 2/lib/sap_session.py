# sap_session.py
# lib/sap_session.py
# filepath: c:\Users\WRL1PO\Documents\Projeto_Inventario\lib\sap_session.py
from playwright.sync_api import Playwright, Browser, BrowserContext, Page
from .config import settings
from .logger import get_logger
from .wait_utils import wait_page_idle
from .page_actions import (
    fill_role_textbox,
    ensure_post_action_stable,
)
from . import selectors
import time
import re  # <-- adicionado

log = get_logger("sap")

class SAPSession:
    def __init__(self, playwright: Playwright):
        self.playwright = playwright
        self.browser: Browser | None = None
        self.context: BrowserContext | None = None
        self.page: Page | None = None
        self._system_message_handled: bool = False  # controle para executar só uma vez

    def start(self):
        log.info("Iniciando browser.")
        self.browser = self.playwright.chromium.launch(
            headless=settings.HEADLESS,
            slow_mo=settings.SLOW_MO
        )
        self.context = self.browser.new_context()
        self.page = self.context.new_page()
        return self

    def goto_base(self):
        log.info(f"Acessando URL: {settings.BASE_URL}")
        self.page.goto(settings.BASE_URL, wait_until="load")
        wait_page_idle(self.page)
        self._try_dismiss_initial_system_message()  # nova chamada

    def open_transaction(self, code: str):
        log.info(f"Abrindo transação: {code}")
        fill_role_textbox(self.page, selectors.TX_INPUT_ROLE, code, press_enter=True)
        ensure_post_action_stable(self.page)

    def set_inventory_number(self, number: str):
        fill_role_textbox(self.page, ("textbox", "Warehouse Number / Warehouse"), "BR2", press_enter=False)
        log.info("Warehouse Number / Warehouse preenchido com 'BR2'.")
        time.sleep(1)
        log.info(f"Definindo inventário: {number}")
        fill_role_textbox(self.page, selectors.INVENTORY_NUMBER_ROLE, number, press_enter=True)

    def close(self):
        log.info("Encerrando sessão.")
        try:
            if self.context:
                self.context.close()
            if self.browser:
                self.browser.close()
        except Exception as e:
            log.error(f"Erro ao fechar sessão: {e}")

    def _try_dismiss_initial_system_message(self):
        """
        Se na primeira abertura aparecer a mensagem com todas as palavras-chave,
        envia ESC (ou clica no botão 'Cancel (Escape)').
        Só executa uma vez.
        """
        if self._system_message_handled:
            return
        keywords = [
            "System Messages",
            "Author",
            "Message Text",
            "System Copy Refresh",
            "Data copied from",
            "Time stamp",
        ]
        try:
            html_lower = self.page.content().lower()
            if all(k.lower() in html_lower for k in keywords):
                log.info("Mensagem inicial 'System Copy Refresh' detectada. Tentando fechar via ESC.")
                try:
                    btn = self.page.get_by_title(re.compile(r"Cancel \(Escape\)", re.I))
                    if btn.count() > 0 and btn.first.is_visible():
                        btn.first.click()
                        log.info("Clique em 'Cancel (Escape)'.")
                    else:
                        self.page.keyboard.press("Escape")
                        log.info("Tecla Escape enviada.")
                except Exception as e:
                    log.warning(f"Falha ao clicar em 'Cancel (Escape)': {e}. Tentando Escape direto.")
                    try:
                        self.page.keyboard.press("Escape")
                    except Exception:
                        pass
                self._system_message_handled = True
            else:
                log.debug("Mensagem inicial especial não detectada.")
        except Exception as e:
            log.debug(f"Não foi possível verificar mensagem inicial: {e}")