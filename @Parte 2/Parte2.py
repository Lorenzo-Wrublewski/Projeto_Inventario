# parte2.py
import sys
import time
import csv
from pathlib import Path
from playwright.sync_api import sync_playwright
from lib.logger import get_logger
from lib.config import settings
from lib.sap_session import SAPSession
from lib.exceptions import (
    ElementNotFound,
    ActionTimeout,
    PageNotIdle,
    SAPMessageError,
    AutomationError,
)
from lib.error_handling import handle_flow_exception
from lib.single_record_entry import process_single_record_entries

log = get_logger("main")

SINGLE_RECORD_FILE = r"C:\Users\WRL1PO\Documents\relatorio2.xlsb"
REFERENCE_REPORT_FILE = r"C:\Users\WRL1PO\Downloads\Template_LX03_test.xlsx"

def run(transaction_code: str = "LI11N", inventory_number: str | None = None):
    with sync_playwright() as pw:
        sap = SAPSession(pw).start()
        try:
            try:
                sap.goto_base()
            except (ActionTimeout, PageNotIdle) as e:
                handle_flow_exception(e, sap, "goto_base")
                raise
            try:
                sap.open_transaction(transaction_code)
            except (ElementNotFound, ActionTimeout) as e:
                handle_flow_exception(e, sap, "open_transaction")
                raise

            data_path = Path(SINGLE_RECORD_FILE)
            ref_path = Path(REFERENCE_REPORT_FILE)

            log.info(f"Iniciando processamento Single Record Entry via arquivo: {data_path}")
            if not data_path.is_file():
                log.error(f"Arquivo não encontrado: {data_path} (pulando etapa)")
            else:
                try:
                    try:
                        reference_exists = False
                        try:
                            reference_exists = ref_path.is_file()
                        except PermissionError:
                            log.warning(f"Sem permissão para acessar relatório referência: {ref_path}. Ignorando comparação UD.")
                        process_single_record_entries(
                            sap.page,
                            str(data_path),
                            reference_report_path=str(ref_path) if reference_exists else None
                        )
                    except PermissionError as e:
                        log.warning(f"Permissão negada ao abrir relatório referência: {e}. Prosseguindo sem comparação UD.")
                        process_single_record_entries(
                            sap.page,
                            str(data_path),
                            reference_report_path=None
                        )
                except Exception as e:
                    handle_flow_exception(e, sap, "single_record_entries")
                    raise

            log.info("Fluxo concluído com sucesso.")

            if settings.FINAL_PAUSE_S > 0:
                log.info(f"Pausa final de {settings.FINAL_PAUSE_S}s para visualização.")
                time.sleep(120)  # mantendo override

            if settings.REQUIRE_KEYPRESS_END:
                try:
                    input("Pressione ENTER para encerrar o navegador...")
                except Exception:
                    pass

        except SAPMessageError as e:
            handle_flow_exception(e, sap, "sap_message")
            raise
        except AutomationError as e:
            handle_flow_exception(e, sap, "automation_generic")
            raise
        except Exception as e:
            handle_flow_exception(e, sap, "unexpected")
            raise
        finally:
            sap.close()


if __name__ == "__main__":
    tx = sys.argv[1] if len(sys.argv) > 1 else "LI11N"
    inv_arg = sys.argv[2] if len(sys.argv) > 2 else None
    run(tx, inv_arg)
