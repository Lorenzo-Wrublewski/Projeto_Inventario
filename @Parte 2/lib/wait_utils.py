# wait_utils.py
# lib/wait_utils.py
# filepath: c:\Users\WRL1PO\Documents\Projeto_Inventario\lib\wait_utils.py
import time
from typing import Callable, Optional
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeout
from .config import settings
from .logger import get_logger
from .exceptions import ActionTimeout, PageNotIdle, ElementNotFound

log = get_logger("wait")

def wait_for(
    predicate: Callable[[], bool],
    timeout_ms: Optional[int] = None,
    interval: Optional[float] = None,
    action_desc: str = ""
):
    timeout_ms = timeout_ms or settings.DEFAULT_TIMEOUT
    interval = interval or settings.WAIT_POLL_INTERVAL
    start = time.time()
    while True:
        try:
            if predicate():
                if action_desc:
                    log.info(f"OK: {action_desc}")
                return True
        except Exception:
            pass
        elapsed_ms = (time.time() - start) * 1000
        if elapsed_ms > timeout_ms:
            raise ActionTimeout(
                f"Timeout após {int(elapsed_ms)}ms aguardando: {action_desc or predicate.__name__}",
                context=action_desc or predicate.__name__
            )
        time.sleep(interval)

def wait_for_locator_visible(page: Page, locator_str: str, timeout_ms: Optional[int] = None):
    desc = f"visibilidade de '{locator_str}'"
    log.info(f"Aguardando {desc}")
    try:
        page.wait_for_selector(locator_str, state="visible", timeout=timeout_ms or settings.DEFAULT_TIMEOUT)
        return page.locator(locator_str)
    except PlaywrightTimeout:
        raise ElementNotFound(f"Elemento não visível: {locator_str}", context=locator_str)

def wait_page_idle(page: Page, timeout_ms: Optional[int] = None):
    timeout_ms = timeout_ms or settings.PAGE_IDLE_TIMEOUT
    start = time.time()
    page.wait_for_load_state("load", timeout=timeout_ms)
    try:
        page.wait_for_load_state("networkidle", timeout=3000)
    except Exception:
        pass
    last_html_len = -1
    stable_loops = 0
    while (time.time() - start) * 1000 < timeout_ms:
        html = page.content()
        curr_len = len(html)
        if curr_len == last_html_len:
            stable_loops += 1
        else:
            stable_loops = 0
        last_html_len = curr_len
        if stable_loops >= 2:
            log.info("Página estável.")
            return
        time.sleep(0.4)
    elapsed = int((time.time() - start) * 1000)
    raise PageNotIdle(f"Página não ficou estável em {elapsed}ms (timeout {timeout_ms}ms).", context="wait_page_idle")