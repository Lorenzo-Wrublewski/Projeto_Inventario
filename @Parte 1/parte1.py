# parte1.py
# c:\Users\WRL1PO\Documents\Projeto_Inventario\@Parte 1\parte1.py
import sys
from lib.utils import run_main
from lib.logger import log

def main():
    log("Iniciando automação Parte 1")
    run_main()
    log("Finalizado Parte 1")

if __name__ == "__main__":
    if "" not in sys.path:
        sys.path.append("")
    main()