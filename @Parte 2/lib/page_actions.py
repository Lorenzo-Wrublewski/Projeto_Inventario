# page_actions.py
# lib/page_actions.py
# filepath: c:\Users\WRL1PO\Documents\Projeto_Inventario\lib\page_actions.py
import time
from typing import Optional, Tuple
from playwright.sync_api import Page
from .config import settings
from .logger import get_logger
from .wait_utils import wait_for_locator_visible, wait_for
from . import selectors

log = get_logger("actions")

def _post_action_delay(step_desc: str):
    if settings.VERBOSE_STEPS:
        log.info(f"[STEP] {step_desc}")
    if settings.ACTION_DELAY_MS > 0:
        time.sleep(settings.ACTION_DELAY_MS / 1000.0)

def _role_locator(page: Page, role: str, name: str):
    return page.get_by_role(role, name=name)

def fill_role_textbox(page: Page, role_and_name: Tuple[str, str], value: str, press_enter: bool = False):
    role, name = role_and_name
    desc = f"Preencher '{name}' com '{value}'"
    log.info(desc)
    locator = _role_locator(page, role, name)
    locator.wait_for(state="visible", timeout=settings.DEFAULT_TIMEOUT)
    locator.click()
    locator.fill(value)
    if press_enter:
        locator.press("Enter")
    _post_action_delay(desc)

def press_enter_role(page: Page, role_and_name: Tuple[str, str]):
    role, name = role_and_name
    desc = f"Pressionar Enter em '{name}'"
    log.info(desc)
    locator = _role_locator(page, role, name)
    locator.wait_for(state="visible", timeout=settings.DEFAULT_TIMEOUT)
    locator.press("Enter")
    _post_action_delay(desc)

def safe_press(page: Page, key: str, desc: str = ""):
    step = f"Press key {key} {desc}".strip()
    log.info(step)
    page.keyboard.press(key)
    _post_action_delay(step)

def click_when_visible(page: Page, selector: str, desc: str = ""):
    loc = wait_for_locator_visible(page, selector)
    step = f"Clique em {desc or selector}"
    log.info(step)
    loc.click()
    _post_action_delay(step)

def handle_popups_if_any(page: Page):
    dialogs = page.locator(selectors.POPUP_DIALOG_SELECTOR)
    if dialogs.count() > 0:
        step = "Popup detectado. Tentando fechar."
        log.info(step)
        buttons = page.locator(selectors.POPUP_OK_BUTTONS)
        if buttons.count() > 0:
            try:
                buttons.first.click(timeout=2000)
                log.info("Popup fechado.")
            except Exception:
                pass
        _post_action_delay(step)

def read_status_message(page: Page) -> Optional[str]:
    try:
        loc = page.locator(selectors.STATUS_BAR_SELECTOR)
        if loc.count() > 0:
            text = loc.first.inner_text().strip()
            return text or None
    except Exception:
        return None
    return None

def wait_status_clear(page: Page, timeout_ms: Optional[int] = None):
    keywords = ("Processando", "Carregando", "Loading", "Aguarde")
    def _clear():
        msg = read_status_message(page)
        if not msg:
            return True
        lowered = msg.lower()
        return not any(k.lower() in lowered for k in keywords)
    wait_for(_clear, timeout_ms=timeout_ms, action_desc="Status livre")

def ensure_post_action_stable(page: Page):
    handle_popups_if_any(page)
    wait_status_clear(page)
    _post_action_delay("Estado estável pós-ação")