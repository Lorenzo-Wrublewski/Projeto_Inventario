# error_handling.py
# lib/error_handling.py
# filepath: c:\Users\WRL1PO\Documents\Projeto_Inventario\lib\error_handling.py
import os
from datetime import datetime
from typing import Tuple, Type
from .logger import get_logger
from .config import settings
from .exceptions import (
    AutomationError,
    ElementNotFound,
    ActionTimeout,
    PageNotIdle,
    SAPMessageError,
)

log = get_logger("errors")

_ERROR_MAP: dict[Type[Exception], Tuple[str, str]] = {
    ElementNotFound: ("ELEMENTO", "Verifique seletor / papel / nome."),
    ActionTimeout: ("TIMEOUT", "Aumente timeout ou valide condição."),
    PageNotIdle: ("PAGE_IDLE", "Talvez carregamento prolongado / redes."),
    SAPMessageError: ("SAP_MSG", "Erro funcional retornado pelo SAP."),
    AutomationError: ("AUTO_ERR", "Erro genérico de automação."),
}

def _screenshot(page, prefix: str):
    if not settings.SCREENSHOT_ON_ERROR or page is None:
        return
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    name = f"{prefix}_{ts}.png"
    try:
        page.screenshot(path=name)
        log.info(f"Screenshot salva: {name}")
    except Exception as e:
        log.error(f"Falha ao salvar screenshot: {e}")

def handle_flow_exception(exc: Exception, sap_session, stage: str):
    """
    Centraliza tratamento: log estruturado + screenshot.
    Retorna código simbólico do erro.
    """
    page = getattr(sap_session, "page", None)
    code = "UNKNOWN"
    help_text = "Sem dica."

    for etype, (c, help_) in _ERROR_MAP.items():
        if isinstance(exc, etype):
            code = c
            help_text = help_
            break

    log.error(f"[ERRO][{code}] Stage={stage} | {exc} | Hint: {help_text}")
    _screenshot(page, f"err_{code.lower()}")
    return code