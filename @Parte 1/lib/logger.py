# logger.py
# c:\Users\WRL1PO\Documents\Projeto_Inventario\@Parte 1\lib\logger.py
import datetime
import sys
from typing import Any

def _ts() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def log(*msg: Any, level: str = "INFO") -> None:
    text = " ".join(str(m) for m in msg)
    sys.stdout.write(f"[{_ts()}][{level}] {text}\n")
    sys.stdout.flush()