# storages.py
# c:\Users\WRL1PO\Documents\Projeto_Inventario\@Parte 1\lib\storages.py
import os
from .logger import log

def load_storages(csv_path: str) -> list[str]:
    """
    Lê o arquivo CSV contendo códigos de Storage Type.
    Suporta separação por ; ou linha. Ignora vazios.
    """
    if not os.path.isfile(csv_path):
        log(f"Arquivo de storages não encontrado: {csv_path}", level="WARN")
        return []
    storages: list[str] = []
    with open(csv_path, "r", encoding="utf-8") as f:
        raw = f.read()
    # Troca quebras de linha por ponto e vírgula para unificar
    raw = raw.replace("\r", "")
    # Separadores: ; ou \n
    tokens = []
    for part in raw.split("\n"):
        tokens.extend(part.split(";"))
    for t in tokens:
        code = t.strip().upper()
        if code:
            storages.append(code)
    # Remove duplicados mantendo ordem
    seen = set()
    ordered = []
    for s in storages:
        if s not in seen:
            seen.add(s)
            ordered.append(s)
    log(f"Storages carregados: {ordered}")
    return ordered