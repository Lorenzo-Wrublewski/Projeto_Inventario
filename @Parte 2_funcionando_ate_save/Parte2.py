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
# NOVOS IMPORTS
import pyodbc
import pandas as pd
import unicodedata

def _norm(v: str | None) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return s.upper()

def fetch_counting_records(reference_report_path: str) -> list[dict]:
    """
    Consulta banco (Tipo Deposito='H0A') e associa DOC vindo de Template_RPA.
    Busca não sequencial: cada registro do banco procura em qualquer linha do template.
    Índices usados:
      full: (CENTRO, DEPOSITO, MATERIAL, BIN)
      mat_bin: (MATERIAL, BIN)
      bin_only: BIN
    Só retorna registros com DOC e campos Material/Centro/Depósito não vazios.
    """
    log.info("Consultando view dbo.vw_PowerBI_DataTable (Tipo Deposito='H0A')...")
    sql = """
        SELECT
            Centro,
            Deposito,
            [Posição no Deposito] AS PosicaoDeposito,
            Material,
            [Tipo Deposito] AS TipoDeposito,
            [Quantidade Eleita] AS QuantidadeEleita
        FROM dbo.vw_PowerBI_DataTable
        WHERE [Tipo Deposito] = 'H0A'
    """
    import pandas as pd
    try:
        with pyodbc.connect(DB_CONNECTION_STRING) as conn:
            df_db = pd.read_sql(sql, conn)
    except Exception as e:
        log.error(f"Falha consulta banco: {e}")
        raise
    log.info(f"Registros banco (H0A): {len(df_db)}")

    log.info(f"Lendo template estoque: {reference_report_path}")
    try:
        df_tpl = pd.read_excel(reference_report_path, engine="openpyxl")
    except Exception as e:
        log.error(f"Erro lendo template: {e}")
        raise

    # Renomear colunas para padrão
    rename_map = {
        "Depósito": "Deposito",
        "Posição no depósito": "PosicaoDeposito",
        "Tipo de depósito": "TipoDeposito",
    }
    for k, v in rename_map.items():
        if k in df_tpl.columns:
            df_tpl = df_tpl.rename(columns={k: v})

    # Garantir colunas
    for c in ["Centro","Deposito","PosicaoDeposito","Material","DOC","TipoDeposito"]:
        if c not in df_tpl.columns:
            df_tpl[c] = ""
        df_tpl[c] = df_tpl[c].fillna("").astype(str).str.strip()

    # Criar índices ignorando linhas sem DOC
    full_index: dict[tuple, str] = {}
    mat_bin_index: dict[tuple, str] = {}
    bin_index: dict[str, str] = {}

    for _, row in df_tpl.iterrows():
        doc = row.get("DOC","").strip()
        if not doc:
            continue
        centro = _norm(row.get("Centro"))
        deposito = _norm(row.get("Deposito"))
        material = _norm(row.get("Material"))
        bin_ = _norm(row.get("PosicaoDeposito"))
        if centro and deposito and material and bin_:
            full_index[(centro, deposito, material, bin_)] = doc
        elif material and bin_:
            mat_bin_index[(material, bin_)] = doc
        elif bin_:
            bin_index[bin_] = doc

    log.info(f"Índice DOC: full={len(full_index)} mat_bin={len(mat_bin_index)} bin={len(bin_index)}")

    # Normalizar banco
    for c in ["Centro","Deposito","PosicaoDeposito","Material"]:
        df_db[c] = df_db[c].fillna("").astype(str).str.strip()

    records: list[dict] = []
    hit_full = hit_mat_bin = hit_bin = 0

    for _, r in df_db.iterrows():
        centro_n = _norm(r["Centro"])
        deposito_n = _norm(r["Deposito"])
        material_n = _norm(r["Material"])
        bin_n = _norm(r["PosicaoDeposito"])

        doc = None
        key_full = (centro_n, deposito_n, material_n, bin_n)
        if key_full in full_index:
            doc = full_index[key_full]; hit_full += 1
        elif (material_n, bin_n) in mat_bin_index:
            doc = mat_bin_index[(material_n, bin_n)]; hit_mat_bin += 1
        elif bin_n in bin_index:
            doc = bin_index[bin_n]; hit_bin += 1

        if not doc:
            continue  # sem DOC não lança
        # Exigir campos obrigatórios preenchidos
        if not (r["Material"] and r["Centro"] and r["Deposito"]):
            continue

        records.append({
            "center": r["Centro"],
            "deposit": r["Deposito"],
            "bin": r["PosicaoDeposito"],
            "material": r["Material"],
            "quantity": r.get("QuantidadeEleita"),
            "doc": doc,
            "deposit_type": r.get("TipoDeposito")
        })

    log.info(f"Associados DOC: full={hit_full} mat_bin={hit_mat_bin} bin={hit_bin} | Final={len(records)}")
    if not records:
        log.warning("Nenhum registro associado. Verifique se DOC corresponde às chaves (Material/Centro/Depósito/Bin).")
    else:
        log.info(f"Exemplo registros prontos: {records[:5]}")
    return records

REFERENCE_REPORT_FILE = r"\\ca0vm0126\FONTES\Python\Projeto_Inventário\Template_RPA.xlsx"

DB_CONNECTION_STRING = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=CA0DEVSQL.br.bosch.com;"
    "DATABASE=InventoryManagementPoPHomolog;"
    "UID=InventoryManagementPoPHomolog;"
    "PWD=_PHc5JlWkIMPH;"
    "Encrypt=no;"
    "TrustServerCertificate=yes;"
)

log = get_logger("main")

def run(transaction_code: str = "LI11N", inventory_number: str | None = None):
    with sync_playwright() as pw:
        sap = SAPSession(pw).start()
        try:
            sap.goto_base()
            sap.open_transaction(transaction_code)

            ref_path = Path(REFERENCE_REPORT_FILE)
            if not ref_path.is_file():
                log.error(f"Template não encontrado: {ref_path}")
                raise FileNotFoundError(str(ref_path))

            try:
                records = fetch_counting_records(str(ref_path))
            except Exception as e:
                handle_flow_exception(e, sap, "fetch_counting_records")
                raise

            if not records:
                log.warning("Nenhum registro retornado do banco. Encerrando.")
            else:
                try:
                    process_single_record_entries(
                        sap.page,
                        reference_report_path=str(ref_path),
                        records=records
                    )
                except Exception as e:
                    handle_flow_exception(e, sap, "single_record_entries")
                    raise

            log.info("Fluxo concluído com sucesso.")

            if settings.FINAL_PAUSE_S > 0:
                log.info(f"Pausa final {settings.FINAL_PAUSE_S}s.")
                time.sleep(settings.FINAL_PAUSE_S)

            if settings.REQUIRE_KEYPRESS_END:
                try:
                    input("ENTER para fechar...")
                except Exception:
                    pass

        except SAPMessageError as e:
            handle_flow_exception(e, sap, "sap_message"); raise
        except AutomationError as e:
            handle_flow_exception(e, sap, "automation_generic"); raise
        except Exception as e:
            handle_flow_exception(e, sap, "unexpected"); raise
        finally:
            sap.close()

if __name__ == "__main__":
    tx = sys.argv[1] if len(sys.argv) > 1 else "LI11N"
    inv_arg = sys.argv[2] if len(sys.argv) > 2 else None
    run(tx, inv_arg)
