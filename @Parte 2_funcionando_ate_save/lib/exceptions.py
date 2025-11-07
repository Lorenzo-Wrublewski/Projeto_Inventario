# exceptions.py
# lib/exceptions.py
# filepath: c:\Users\WRL1PO\Documents\Projeto_Inventario\lib\exceptions.py
from typing import Optional

class AutomationError(Exception):
    """
    Base para exceções específicas da automação.
    Armazena mensagem base + contexto opcional.
    """
    default_message = "Erro de automação."

    def __init__(self, message: Optional[str] = None, *, context: Optional[str] = None):
        self.context = context
        final_msg = message or self.default_message
        if context:
            final_msg = f"{final_msg} | Contexto: {context}"
        super().__init__(final_msg)

class ElementNotFound(AutomationError):
    default_message = "Elemento não encontrado."

class ActionTimeout(AutomationError):
    default_message = "Timeout ao aguardar ação."

class PageNotIdle(AutomationError):
    default_message = "Página não ficou estável."

class SAPMessageError(AutomationError):
    default_message = "Mensagem de erro retornada pelo SAP."