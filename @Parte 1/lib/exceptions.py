# exceptions.py
# c:\Users\WRL1PO\Documents\Projeto_Inventario\@Parte 1\lib\exceptions.py
class SapAutomationError(Exception):
    """Erro genérico na automação SAP."""

class SapElementNotFound(SapAutomationError):
    """Elemento não encontrado dentro do tempo limite."""

class SapTimeoutError(SapAutomationError):
    """Timeout de espera por condição."""