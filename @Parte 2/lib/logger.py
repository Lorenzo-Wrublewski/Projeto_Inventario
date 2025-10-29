# logger.py
# lib/logger.py
# filepath: c:\Users\WRL1PO\Documents\Projeto_Inventario\lib\logger.py
import logging
import os
from datetime import datetime

_LOGGER = None

def get_logger(name: str = "app"):
    global _LOGGER
    if _LOGGER:
        return _LOGGER

    log_dir = os.path.join(os.getcwd(), "logs")
    os.makedirs(log_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d")
    file_path = os.path.join(log_dir, f"run_{timestamp}.log")

    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)

    fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(name)s | %(message)s")

    sh = logging.StreamHandler()
    sh.setFormatter(fmt)
    logger.addHandler(sh)

    fh = logging.FileHandler(file_path, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    _LOGGER = logger
    return logger