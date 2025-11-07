# selectors.py
# lib/selectors.py
# filepath: c:\Users\WRL1PO\Documents\Projeto_Inventario\lib\selectors.py
TX_INPUT_ROLE = ("textbox", "Enter transaction code")
INVENTORY_NUMBER_ROLE = ("textbox", "Number of system inventory")

STATUS_BAR_SELECTOR = "div[id*='statusbar'], span[id*='status']"
POPUP_DIALOG_SELECTOR = "div[role='dialog'], div[aria-modal='true']"
POPUP_OK_BUTTONS = "button:has-text('OK'), button:has-text('Ok'), button:has-text('Continuar')"
ERROR_MESSAGE_SELECTOR = "span[class*='sapMMsgStripError'], .msg-error"