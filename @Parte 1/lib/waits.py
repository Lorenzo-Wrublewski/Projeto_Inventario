# waits.py
# c:\Users\WRL1PO\Documents\Projeto_Inventario\@Parte 1\lib\waits.py
import time
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError
from .logger import log
from .exceptions import SapElementNotFound, SapTimeoutError

GLOBAL_ACTION_DELAY: float = 0.0

def set_global_action_delay(seconds: float):
    global GLOBAL_ACTION_DELAY
    GLOBAL_ACTION_DELAY = max(0.0, seconds)
    log(f"Atraso global entre ações definido: {GLOBAL_ACTION_DELAY:.2f}s")

def wait_for_locator_visible(page: Page, selector, timeout: float = 10.0, description: str = ""):
    try:
        loc = selector if hasattr(selector, "wait_for") else page.locator(selector)
        loc.wait_for(state="visible", timeout=timeout * 1000)
        if description:
            log(f"Elemento visível: {description}")
        return loc
    except PlaywrightTimeoutError:
        raise SapElementNotFound(f"Elemento não visível (timeout {timeout}s): {description or selector}")

def wait_until_any(page: Page, locators, timeout: float, poll: float = 0.5):
    end = time.time() + timeout
    while time.time() < end:
        for desc, locator in locators:
            try:
                if locator.is_visible():
                    log(f"Detectado: {desc}")
                    return desc, locator
            except Exception:
                pass
        time.sleep(poll)
    raise SapTimeoutError("Nenhum dos elementos esperados apareceu dentro do tempo limite.")

def safe_click(locator, description: str = "", timeout: float = 10.0):
    locator.wait_for(state="visible", timeout=timeout * 1000)
    locator.click()
    if description:
        log(f"Click: {description}")
    _apply_global_delay("Após click")

def safe_fill(locator, value: str, description: str = "", delay: float = 0.0):
    locator.wait_for(state="visible")
    locator.fill(value)
    if delay:
        time.sleep(delay)
    if description:
        log(f"Preenchido '{value}' em {description}")
    _apply_global_delay("Após fill")

def wait_seconds(seconds: float, reason: str = ""):
    if seconds <= 0:
        return
    if reason:
        log(f"Aguardando {seconds:.1f}s ({reason})")
    time.sleep(seconds)

def wait_for_no_busy(page: Page, timeout: float, min_stable: float = 1.0, poll: float = 0.3):
    busy_selectors = [
        "div.sapUiBusy",
        "div[class*='BusyIndicator']",
        "div[id*='busy']",
        "div[class*='urMsgBarInProgress']",
        "img[alt*='Working']",
        "img[alt*='Carregando']"
    ]
    end = time.time() + timeout
    stable_start = None
    while time.time() < end:
        any_busy = False
        for sel in busy_selectors:
            try:
                if page.locator(sel).first.is_visible():
                    any_busy = True
                    break
            except Exception:
                pass
        if any_busy:
            stable_start = None
        else:
            if stable_start is None:
                stable_start = time.time()
            if time.time() - stable_start >= min_stable:
                log("Interface estável (sem busy).")
                return True
        time.sleep(poll)
    log("Timeout aguardando estabilização (busy persistente).", level="WARN")
    return False

def wait_for_interface_stable(page: Page, timeout: float, min_stable: float):
    log(f"Iniciando verificação de estabilidade da interface (timeout {timeout}s).")
    wait_for_no_busy(page, timeout=timeout, min_stable=min_stable)

def _apply_global_delay(reason: str):
    if GLOBAL_ACTION_DELAY > 0:
        wait_seconds(GLOBAL_ACTION_DELAY, f"Delay global - {reason}")