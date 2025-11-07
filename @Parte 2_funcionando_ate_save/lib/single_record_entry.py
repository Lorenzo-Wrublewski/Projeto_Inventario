# single_record_entry.py
# filepath: c:\Users\WRL1PO\Documents\Projeto_Inventario\@Parte 2\lib\single_record_entry.py
import csv
import re
import time
import unicodedata
from pathlib import Path
from typing import List, Dict, Optional, Tuple
from playwright.sync_api import Page
from .logger import get_logger
from .exceptions import ElementNotFound
from .config import settings
from .wait_utils import wait_for  # reutiliza função genérica
from .page_actions import fill_role_textbox  # se ainda não importado

log = get_logger("single_record")

# Campos padrão usados no SAP + novos para lógica UD
EXPECTED_HEADERS = {
    "storage bin": "storage_bin",
    "material number": "material_number",
    "counted quantity in alternative unit of measure": "quantity_alt",
    "storage location": "storage_location",
    "plant": "plant",
}

# Aliases ampliados
SRE_HEADER_ALIASES = {
    "posicao no deposito": "storage_bin",
    "posição no depósito": "storage_bin",
    "material": "material_number",
    "qtd.contada": "counted_quantity",  # campo lógico principal de contagem
    "deposito": "storage_location",
    "depósito": "storage_location",
    "centro": "plant",
    "tipo de depósito": "storage_type",
    "tipo de deposito": "storage_type",
    "estoque total": "stock_total",
    "ud": "ud",
}

# Documento inventário
SRE_HEADER_ALIASES.update({
    "documento inventario": "inventory_record",
    "documento inventário": "inventory_record",
})

def _norm(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").lower().strip()

def _parse_number(num_str: str) -> float:
    """
    Converte '326,00' ou '5,00' em float 326.00 / 5.00.
    Ignora vazio => 0.0.
    """
    if not num_str:
        return 0.0
    s = str(num_str).strip()
    if not s:
        return 0.0
    # remove possíveis separadores de milhar (.)
    if s.count(",") == 1 and s.count(".") >= 1:
        # heurística: remover pontos e trocar vírgula
        s = s.replace(".", "")
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

def load_single_record_csv(csv_path: str, delimiter: str = ";") -> List[Dict[str, str]]:
    path = Path(csv_path)
    if not path.is_file():
        raise FileNotFoundError(f"Arquivo CSV não encontrado: {csv_path}")
    rows: List[Dict[str, str]] = []
    with path.open("r", encoding="utf-8-sig") as f:
        reader = csv.reader(f, delimiter=delimiter)
        header_map: Dict[int, str] = {}
        try:
            raw_header = next(reader)
        except StopIteration:
            return rows
        for idx, col in enumerate(raw_header):
            key_norm = _norm(col)
            if key_norm in SRE_HEADER_ALIASES:
                header_map[idx] = SRE_HEADER_ALIASES[key_norm]
            elif key_norm in EXPECTED_HEADERS:
                header_map[idx] = EXPECTED_HEADERS[key_norm]
        if not header_map:
            log.warning("Cabeçalho não reconhecido no CSV principal.")
        for line in reader:
            if not any(cell.strip() for cell in line):
                continue
            rows.append({v: (line[i].strip() if i < len(line) else "") for i, v in header_map.items()})
    log.info(f"Carregado CSV '{csv_path}' com {len(rows)} registros.")
    return rows

def _fix_storage_bin(val: str) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    if not s:
        return ""
    if re.fullmatch(r"\d+(\.0)?", s):
        if s.endswith(".0"):
            s = s[:-2]
        if 5 <= len(s) < 10:
            s = s.zfill(10)
        return s
    return s

def load_single_record_excel(excel_path: str) -> List[Dict[str, str]]:
    import pandas as pd
    if not Path(excel_path).is_file():
        raise FileNotFoundError(f"Arquivo Excel não encontrado: {excel_path}")
    df = pd.read_excel(excel_path, engine="pyxlsb")
    col_map: Dict[str, str] = {}
    for col in df.columns:
        n = _norm(str(col))
        if n in SRE_HEADER_ALIASES:
            col_map[col] = SRE_HEADER_ALIASES[n]
        elif n in EXPECTED_HEADERS:
            col_map[col] = EXPECTED_HEADERS[n]
    if "inventory_record" not in col_map.values():
        raise ValueError("Coluna 'Documento inventário' obrigatória não encontrada no Excel.")
    rows: List[Dict[str, str]] = []
    for _, r in df.iterrows():
        std = {
            "inventory_record": "",
            "storage_bin": "",
            "material_number": "",
            "quantity_alt": "",
            "counted_quantity": "",
            "storage_location": "",
            "plant": "",
            "storage_type": "",
            "stock_total": "",
            "ud": ""
        }
        for orig, std_key in col_map.items():
            std[std_key] = str(r.get(orig, "")).strip()
        std["storage_bin"] = _fix_storage_bin(std["storage_bin"])
        # Se counted_quantity presente, replica para quantity_alt para envio ao SAP
        if std.get("counted_quantity"):
            std["quantity_alt"] = std["counted_quantity"]
        # Filtra linhas essencialmente vazias
        if any(v for v in std.values()):
            rows.append(std)
    log.info(f"Carregado Excel '{excel_path}' com {len(rows)} registros.")
    return rows

def _single_record_button_locator(page: Page):
    time.sleep(1)
    return page.locator("div").filter(has_text=re.compile(r"^Single Record Entry$"))

def load_single_record_file(path: str) -> List[Dict[str, str]]:
    ext = Path(path).suffix.lower()
    if ext == ".xlsb":
        return load_single_record_excel(path)
    if ext == ".csv":
        return load_single_record_csv(path)
    raise ValueError(f"Extensão não suportada: {ext} (use .xlsb ou .csv)")

# Substituir antiga load_marcelo_report por função genérica:
def load_comparison_report(path_str: str) -> List[Dict[str, str]]:
    """
    Lê relatório de comparação (Excel .xlsx/.xls ou .csv).
    Campos relevantes: material_number, storageBin, stock_total, (plant, storage_location se existirem).
    """
    if not path_str:
        return []
    path = Path(path_str)
    try:
        if not path.is_file():
            log.warning(f"Relatório referência não encontrado: {path_str}")
            return []
    except PermissionError:
        log.warning(f"Sem permissão para acessar '{path_str}'. Comparação UD ignorada.")
        return []

    ext = path.suffix.lower()
    rows: List[Dict[str, str]] = []
    if ext in [".xlsx", ".xls"]:
        try:
            import pandas as pd
        except ImportError:
            log.warning("pandas não disponível para leitura do Excel de referência. Ignorando comparação.")
            return []
        try:
            df = pd.read_excel(path_str, engine="openpyxl")
        except PermissionError:
            log.warning(f"Permissão negada ao ler '{path_str}'. Ignorando comparação.")
            return []
        # Mapear colunas
        col_map: Dict[str, str] = {}
        for col in df.columns:
            n = _norm(str(col))
            if n in SRE_HEADER_ALIASES:
                col_map[col] = SRE_HEADER_ALIASES[n]
            elif n in EXPECTED_HEADERS:
                col_map[col] = EXPECTED_HEADERS[n]
        for _, r in df.iterrows():
            entry = {}
            for orig, std_key in col_map.items():
                entry[std_key] = str(r.get(orig, "")).strip()
            if any(entry.values()):
                rows.append(entry)
        log.info(f"Carregado relatório referência Excel '{path_str}' com {len(rows)} registros.")
        return rows

    # CSV fallback
    delimiter = ";"
    try:
        with path.open("r", encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter=delimiter)
            try:
                raw_header = next(reader)
            except StopIteration:
                return rows
            header_map: Dict[int, str] = {}
            for idx, col in enumerate(raw_header):
                n = _norm(col)
                if n in SRE_HEADER_ALIASES:
                    header_map[idx] = SRE_HEADER_ALIASES[n]
                elif n in EXPECTED_HEADERS:
                    header_map[idx] = EXPECTED_HEADERS[n]
            for line in reader:
                if not any(cell.strip() for cell in line):
                    continue
                rows.append({v: (line[i].strip() if i < len(line) else "") for i, v in header_map.items()})
    except PermissionError:
        log.warning(f"Permissão negada ao ler '{path_str}'. Ignorando comparação.")
        return []
    log.info(f"Carregado relatório referência CSV '{path_str}' com {len(rows)} registros.")
    return rows

def _wait_and_click(page: Page, regex_text: str, timeout_ms: Optional[int] = None):
    timeout_ms = timeout_ms or settings.DEFAULT_TIMEOUT
    locator = page.locator("div").filter(has_text=re.compile(regex_text))
    locator.first.wait_for(state="visible", timeout=timeout_ms)
    locator.first.click()

def _fill_field(page: Page, role_name: str, value: str):
    value = value or ""
    tb = page.get_by_role("textbox", name=role_name)
    tb.wait_for(state="visible", timeout=settings.DEFAULT_TIMEOUT)
    tb.click()
    tb.fill(value)

def _click_cancel_once(page: Page):
    """
    Clica em 'Cancel' se visível (não trata confirmação).
    """
    candidates = [
        page.locator("div").filter(has_text=re.compile(r"^Cancel$")),
        page.get_by_role("button", name=re.compile(r"^Cancel$", re.I)),
    ]
    for loc in candidates:
        try:
            if loc.count() > 0 and loc.first.is_visible():
                loc.first.click()
                log.info("Click Cancel")
                time.sleep(0.3)
                return True
        except Exception:
            pass
    log.debug("Botão Cancel não encontrado nessa tentativa.")
    return False

def _confirm_exit_yes(page: Page, timeout_s: float = 3.0):
    """
    Se aparecer o popup de confirmação (Yes/Sim), clica em Yes.
    Procura repetidamente até timeout ou sumir.
    """
    end = time.time() + timeout_s
    clicked = False
    while time.time() < end:
        try:
            # Possíveis variações
            locators = [
                page.get_by_title("Yes"),
                page.get_by_role("button", name=re.compile(r"^(Yes|Sim)$", re.I)),
                page.locator("button").filter(has_text=re.compile(r"^(Yes|Sim)$", re.I)),
                page.locator("div").filter(has_text=re.compile(r"^(Yes|Sim)$", re.I)),
            ]
            for loc in locators:
                if loc.count() > 0 and loc.first.is_visible():
                    loc.first.click()
                    log.info("Click Yes (confirma saída)")
                    clicked = True
                    time.sleep(0.4)
                    # Verifica se ainda existe para nova tentativa
                    break
            if clicked:
                # Checa se o popup sumiu
                still_visible = False
                for loc in locators:
                    try:
                        if loc.count() > 0 and loc.first.is_visible():
                            still_visible = True
                            break
                    except Exception:
                        pass
                if not still_visible:
                    return True
        except Exception:
            pass
        time.sleep(0.2)
    if not clicked:
        log.debug("Popup Yes não detectado (talvez não requerido).")
    else:
        log.warning("Tentativa de confirmar Yes pode não ter fechado popup.")
    return clicked

def _wait_inventory_field(page: Page, timeout_ms: int | None = None):
    timeout_ms = timeout_ms or settings.DEFAULT_TIMEOUT
    start = time.time()
    while (time.time() - start) * 1000 < timeout_ms:
        try:
            fld = page.get_by_role("textbox", name=re.compile(r"Number of system inventory", re.I))
            if fld.count() > 0 and fld.first.is_visible():
                return True
        except Exception:
            pass
        time.sleep(0.25)
    raise ElementNotFound("Campo 'Number of system inventory' não retornou no timeout.")

# --- Lógica de Sequência UD ---

def _detect_ud_sequence(records: List[Dict[str, str]], start_index: int) -> Tuple[int, int]:
    """
    Retorna (inicio, fim_exclusivo) da sequência UD começando em start_index.
    Critérios:
      - UD não vazia
      - Material + storage_bin iguais entre linhas
    """
    first = records[start_index]
    mat = first.get("material_number", "").strip()
    bin_ = first.get("storage_bin", "").strip()
    if not first.get("ud"):
        return (start_index, start_index + 1)
    end = start_index + 1
    while end < len(records):
        r = records[end]
        if not r.get("ud"):
            break
        if r.get("material_number", "").strip() != mat or r.get("storage_bin", "").strip() != bin_:
            break
        end += 1
    return (start_index, end)

def _index_marcelo(records_marcelo: List[Dict[str, str]]) -> Dict[Tuple[str, str], List[Dict[str, str]]]:
    idx: Dict[Tuple[str, str], List[Dict[str, str]]] = {}
    for r in records_marcelo:
        mat = r.get("material_number", "").strip()
        bin_ = r.get("storage_bin", "").strip()
        if not mat or not bin_:
            continue
        key = (mat, bin_)
        idx.setdefault(key, []).append(r)
    return idx

def _reallocate_quantities(original: List[float], case: str) -> List[float]:
    """
    case: 'MENOR', 'MAIOR' ou 'IGUAL'
    Regra 'MENOR': mover menor para primeira posição (rotaciona segmento até índice do menor).
    Regra 'MAIOR': mover maior para última posição.
    IGUAL: retorna original.
    """
    if case == "IGUAL" or len(original) <= 1:
        return original[:]
    new = original[:]
    if case == "MENOR":
        idx_min = min(range(len(original)), key=lambda i: original[i])
        if idx_min != 0:
            moved = original[idx_min]
            # rotaciona esquerda entre [0, idx_min]
            for k in range(idx_min, 0, -1):
                new[k] = original[k - 1]
            new[0] = moved
    elif case == "MAIOR":
        idx_max = max(range(len(original)), key=lambda i: original[i])
        if idx_max != len(original) - 1:
            moved = original[idx_max]
            for k in range(idx_max, len(original) - 1):
                new[k] = original[k + 1]
            new[-1] = moved
    return new

def _format_quantity(val: float) -> str:
    """
    Inteiro sem decimais se parte fracionária zero.
    Caso tenha fração, usa vírgula e remove zeros à direita.
    Ex.: 7.0 -> '7'; 7.50 -> '7,5'; 7.25 -> '7,25'
    """
    if val is None:
        return ""
    try:
        v = float(val)
    except Exception:
        return str(val)
    if v.is_integer():
        return str(int(v))
    s = f"{v:.4f}".rstrip("0").rstrip(".")
    # troca ponto por vírgula para manter convenção local
    s = s.replace(".", ",")
    return s

def _parse_ud_number(ud_str: str) -> int:
    try:
        return int(str(ud_str).strip())
    except Exception:
        return 0

def _adjust_ud_quantities(counted_total: float, ud_rows: List[Dict[str, str]]) -> List[float]:
    """
    Recebe linhas UD (cada com stock_total) e aplica delta conforme regras.
    Retorna lista final de quantidades para cada UD na ordem crescente de UD.
    """
    # Ordena crescente (antigas primeiro)
    ordered = sorted(ud_rows, key=lambda r: _parse_ud_number(r.get("ud", "")))
    original = [_parse_number(r.get("stock_total", "")) for r in ordered]
    soma_ud = sum(original)
    delta = counted_total - soma_ud
    if delta == 0:
        return original  # sem ajuste

    adjusted = original[:]

    if delta < 0:
        # Remover |delta| começando da mais antiga (índice 0 → ...)
        remaining = abs(delta)
        for i in range(len(adjusted)):
            if remaining <= 0:
                break
            can_remove = min(adjusted[i], remaining)
            adjusted[i] -= can_remove
            remaining -= can_remove
    else:
        # Adicionar delta na mais nova (último índice) conforme regra (tudo na última)
        adjusted[-1] += delta

    return adjusted

# NOVO: mover para nível global
def _is_invalid_field(v: str | None) -> bool:
    if v is None:
        return True
    s = str(v).strip().lower()
    return s == "" or s == "nan"

def _save_and_confirm(page: Page, timeout_yes_s: float = 4.0) -> bool:
    """
    Clica em Save e confirma popup Yes/Sim.
    Retorna True se conseguiu salvar e confirmar.
    """
    try:
        save_btn = page.locator("div").filter(has_text=re.compile(r"^Save$"))
        if save_btn.count() == 0:
            save_btn = page.get_by_role("button", name=re.compile(r"^Save$", re.I))
        if save_btn.count() > 0 and save_btn.first.is_visible():
            save_btn.first.click()
            log.info("Click Save (finalizando DOC)")
            time.sleep(0.5)
            _confirm_exit_yes(page, timeout_s=timeout_yes_s)
            return True
        else:
            log.debug("Botão Save não visível na tela atual.")
    except Exception as e:
        log.warning(f"Falha ao clicar Save: {e}")
    return False

def process_single_record_entries(
    page: Page,
    contagem_path: Optional[str] = None,
    reference_report_path: Optional[str] = None,
    records: Optional[List[Dict[str, str]]] = None
):
    """
    Registra apenas se Material, Centro (plant) e Depósito (storage_location) estiverem preenchidos.
    """
    # REMOVIDO: definição interna de _is_invalid_field

    if records is not None:
        filtered = []
        skipped = 0
        for r in records:
            doc = r.get("doc")
            material = r.get("material")
            plant = r.get("center")
            deposit = r.get("deposit")
            if any([
                _is_invalid_field(doc),
                _is_invalid_field(material),
                _is_invalid_field(plant),
                _is_invalid_field(deposit)
            ]):
                skipped += 1
                continue
            filtered.append(r)
        log.info(f"Registros recebidos (DB): {len(records)} | Válidos p/ lançamento: {len(filtered)} | Pulados (campos vazios/nan): {skipped}")
        total = len(filtered)
        # Ordena por DOC preservando ordem original para detectar último de cada DOC
        # Cria lista de mapeados com flag last
        mapped_seq: List[Dict[str, str]] = []
        for idx, r in enumerate(filtered):
            mapped_seq.append({
                "inventory_record": str(r.get("doc")).strip(),
                "material_number": str(r.get("material")).strip(),
                "storage_bin": str(r.get("bin") or "").strip(),
                "plant": str(r.get("center")).strip(),
                "storage_location": str(r.get("deposit")).strip(),
                "storage_type": str(r.get("deposit_type") or "").strip(),
                "counted_quantity": _format_quantity(_parse_number(r.get("quantity"))),
                "quantity_alt": _format_quantity(_parse_number(r.get("quantity"))),
                "ud": "",
                "stock_total": "",
                "__doc__": str(r.get("doc")).strip()
            })
        # Marca último por DOC olhando próximo diferente
        for i in range(len(mapped_seq)):
            cur_doc = mapped_seq[i]["__doc__"]
            next_doc = mapped_seq[i+1]["__doc__"] if i+1 < len(mapped_seq) else None
            mapped_seq[i]["__last_in_doc__"] = (cur_doc != next_doc)

        for idx, rec in enumerate(mapped_seq, start=1):
            _process_single_record(page, rec, idx, total, is_last_in_doc=rec["__last_in_doc__"])
        log.info("Lançamentos (DB) concluídos com lógica Save por DOC.")
        return
    # Preparar registros de contagem
    if records is not None:
        contagem_records: List[Dict[str, str]] = []
        for r in records:
            contagem_records.append({
                "inventory_record": str(r.get("doc") or "").strip(),
                "material_number": str(r.get("material") or "").strip(),
                "storage_bin": str(r.get("bin") or "").strip(),
                "plant": str(r.get("center") or "").strip(),
                "storage_location": str(r.get("deposit") or "").strip(),
                "storage_type": str(r.get("deposit_type") or "").strip(),
                "counted_quantity": str(r.get("quantity") or "0"),  # total contado
                "quantity_alt": str(r.get("quantity") or "0"),
                "ud": "",
                "stock_total": ""  # não vem do DB
            })
        log.info(f"Contagem (DB) carregada: {len(contagem_records)} registros.")
    else:
        if not contagem_path:
            raise ValueError("Forneça 'records' ou 'contagem_path'.")
        contagem_records = load_single_record_file(contagem_path)
        log.info(f"Contagem (arquivo) carregada: {len(contagem_records)} registros.")

    if not contagem_records:
        log.warning("Nenhum registro de contagem disponível.")
        return

    # Carregar template referência
    reference_records: List[Dict[str, str]] = []
    if reference_report_path:
        path = Path(reference_report_path)
        if path.is_file():
            try:
                import pandas as pd
                df_temp = pd.read_excel(reference_report_path, engine="openpyxl")
                # Renomeia conforme mapa usado em Parte2
                rename_map = {
                    'Depósito': 'Deposito',
                    'Posição no depósito': 'Posição no Deposito',
                    'Tipo de depósito': 'TipoDeposito',
                    'Estoque Total': 'EstoqueTotal'
                }
                df_temp = df_temp.rename(columns={c: rename_map[c] for c in df_temp.columns if c in rename_map})
                for _, row in df_temp.iterrows():
                    reference_records.append({
                        "inventory_record": str(row.get("DOC") or "").strip(),
                        "material_number": str(row.get("Material") or "").strip(),
                        "plant": str(row.get("Centro") or "").strip(),
                        "storage_location": str(row.get("Deposito") or "").strip(),
                        "storage_bin": str(row.get("Posição no Deposito") or "").strip(),
                        "storage_type": str(row.get("TipoDeposito") or "").strip(),
                        "stock_total": str(row.get("EstoqueTotal") or "").strip(),
                        "ud": str(row.get("UD") or "").strip()
                    })
                log.info(f"Template referência carregado: {len(reference_records)} linhas.")
            except Exception as e:
                log.warning(f"Falha ao ler template referência: {e}. Prosseguindo sem referência.")
        else:
            log.warning(f"Template referência não encontrado: {reference_report_path}")
    else:
        log.warning("Sem caminho de template referência.")

    if not reference_records:
        # Lançamento direto sem lógica UD
        log.warning("Sem linhas de referência. Lançando registros de contagem diretamente.")
        for idx, rec in enumerate(contagem_records, start=1):
            rec["quantity_alt"] = _format_quantity(_parse_number(rec.get("counted_quantity") or "0"))
            _process_single_record(page, rec, idx, len(contagem_records))
        log.info("Concluído lançamento sem referência.")
        return

    total_contagem = len(contagem_records)
    launch_list: List[Dict[str, str]] = []
    log.info(f"Iniciando composição de lançamentos com referência. Registros contagem={total_contagem}")

    for i, cont in enumerate(contagem_records, start=1):
        material = cont.get("material_number", "").strip()
        plant = cont.get("plant", "").strip()
        storage_location = cont.get("storage_location", "").strip()
        storage_bin = cont.get("storage_bin", "").strip()
        storage_type = cont.get("storage_type", "").strip()
        inv_doc = cont.get("inventory_record", "").strip()
        counted_total = _parse_number(cont.get("counted_quantity") or cont.get("quantity_alt") or "0")

        if not inv_doc:
            log.warning(f"[{i}/{total_contagem}] Sem DOC para material {material}. Ignorado.")
            continue

        # Filtra referência correspondente
        def _match(ref: Dict[str, str]) -> bool:
            if ref.get("material_number", "").strip() != material:
                return False
            if plant and ref.get("plant", "").strip() != plant:
                return False
            if storage_location and ref.get("storage_location", "").strip() != storage_location:
                return False
            if storage_bin and ref.get("storage_bin", "").strip() != storage_bin:
                return False
            # TipoDeposito pode vir mas cont não necessariamente tem
            return True

        matching = [r for r in reference_records if _match(r)]

        if not matching:
            log.info(f"[{i}/{total_contagem}] Sem referência p/ Material={material} Bin={storage_bin}. Lançando total direto.")
            cont["quantity_alt"] = _format_quantity(counted_total)
            launch_list.append(cont)
            continue

        ud_rows = [r for r in matching if r.get("ud")]
        non_ud_rows = [r for r in matching if not r.get("ud")]

        log.info(f"[{i}/{total_contagem}] Material={material} Bin={storage_bin} UDs={len(ud_rows)} SemUD={len(non_ud_rows)} Contado={counted_total}")

        # Distribuição para UDs
        if ud_rows:
            adjusted_values = _adjust_ud_quantities(counted_total, ud_rows)
            ordered_ud = sorted(ud_rows, key=lambda r: _parse_ud_number(r.get("ud")))
            for idx_ud, ud_ref in enumerate(ordered_ud):
                q_final = adjusted_values[idx_ud]
                formatted = _format_quantity(q_final)
                launch_list.append({
                    "inventory_record": inv_doc,
                    "material_number": material,
                    "storage_bin": ud_ref.get("storage_bin", storage_bin),
                    "plant": plant,
                    "storage_location": storage_location,
                    "storage_type": storage_type,
                    "ud": ud_ref.get("ud", ""),
                    "counted_quantity": formatted,
                    "quantity_alt": formatted,
                    "stock_total": ud_ref.get("stock_total", "")
                })
        else:
            # Sem UDs: uma linha com total contado
            cont["quantity_alt"] = _format_quantity(counted_total)
            launch_list.append(cont)

        # Linhas sem UD com estoque total original (opcional)
        for nu in non_ud_rows:
            qty_orig = _parse_number(nu.get("stock_total", "0"))
            formatted = _format_quantity(qty_orig)
            launch_list.append({
                "inventory_record": inv_doc,
                "material_number": material,
                "storage_bin": nu.get("storage_bin", storage_bin),
                "plant": plant,
                "storage_location": storage_location,
                "storage_type": storage_type,
                "ud": "",
                "counted_quantity": formatted,
                "quantity_alt": formatted,
                "stock_total": nu.get("stock_total", "")
            })

    total_launch = len(launch_list)
    log.info(f"Total linhas para lançar: {total_launch}")

    for idx, rec in enumerate(launch_list, start=1):
        _process_single_record(page, rec, idx, total_launch)

    log.info("Lançamentos concluídos com referência.")

def _process_single_record(page: Page, rec: Dict[str, str], idx: int, total: int, seq_info: Optional[str] = None, is_last_in_doc: bool = False):
    if any([
        _is_invalid_field(rec.get("inventory_record")),
        _is_invalid_field(rec.get("material_number")),
        _is_invalid_field(rec.get("plant")),
        _is_invalid_field(rec.get("storage_location"))
    ]):
        log.warning(f"[Registro {idx}/{total}] Campos obrigatórios vazios/nan. Pulado.")
        return
    inv = rec.get("inventory_record", "").strip()
    tag = f"[Registro {idx}/{total}]{'[' + seq_info + ']' if seq_info else ''}"
    try:
        _open_single_record_entry_after_inventory(page, inv)
        log.info(f"{tag} {rec}")

        time.sleep(0.6)
        _fill_field(page, "Storage Bin", rec.get("storage_bin", ""))
        _pause()
        _fill_field(page, "Material Number", rec.get("material_number", ""))
        _pause()
        _fill_field(page, "Counted quantity in alternative unit of measure",
                    _format_quantity(_parse_number(rec.get("quantity_alt") or rec.get("counted_quantity") or "0")))
        _pause()
        qty_zero = rec.get("quantity_alt", "").strip() in ["0", "0.0", "0,0"]
        if qty_zero:
            try:
                page.get_by_text("Zero stock").click()
                log.info(f"{tag} Caixa 'Zero stock' marcada.")
                time.sleep(0.4)
            except Exception:
                log.debug(f"{tag} 'Zero stock' não encontrada.")

        _fill_field(page, "Storage Location", rec.get("storage_location", ""))
        _pause()
        _fill_field(page, "Plant", rec.get("plant", ""))
        _pause()

        # Confirma quantidade (Enter duas vezes)
        try:
            qty_field = page.get_by_role("textbox", name="Counted quantity in alternative unit of measure")
            if qty_field.count() > 0:
                qty_field.first.press("Enter")
                _pause()
                qty_field.first.press("Enter")
                _pause()
        except Exception:
            log.debug(f"{tag} Não conseguiu pressionar Enter no campo quantidade.")

        # Cancelar sequência
        _click_cancel_once(page)
        _pause()

        if is_last_in_doc:
            # Último registro desse DOC: salvar em vez de segundo cancel
            saved = _save_and_confirm(page)
            if not saved:
                log.debug(f"[Registro {idx}/{total}] Save não efetuado, usando Cancel padrão.")
                _click_cancel_once(page)
                _pause()
                _confirm_exit_yes(page, timeout_s=4.0)
            else:
                # Após Save já confirmou Yes; garantir retorno INVENTORY
                try:
                    _wait_inventory_field(page, timeout_ms=settings.DEFAULT_TIMEOUT)
                    log.info("OK: Tela INVENTORY disponível (após Save)")
                except Exception:
                    log.warning("Não confirmou INVENTORY após Save; tentando Yes extra.")
                    _confirm_exit_yes(page, timeout_s=2.0)
        else:
            # Fluxo antigo
            _click_cancel_once(page)
            _pause()
            _confirm_exit_yes(page, timeout_s=4.0)
            try:
                _wait_inventory_field(page, timeout_ms=settings.DEFAULT_TIMEOUT)
                log.info("OK: Tela INVENTORY disponível")
            except Exception:
                log.warning("Inventário não confirmado; tentando Yes extra.")
                _confirm_exit_yes(page, timeout_s=2.0)
    except ElementNotFound as e:
        log.error(f"{tag} Falha elemento: {e}")
    except Exception as e:
        log.error(f"{tag} Erro inesperado: {e}")

WAREHOUSE_VALUE = "BR2"

ACTION_INTERVAL_S = getattr(settings, "SINGLE_RECORD_INTERVAL_S", 0.4)
SHORT_SLEEP = getattr(settings, "SHORT_SLEEP_S", 0.25)

def _pause(tag: str = ""):
    time.sleep(ACTION_INTERVAL_S)

def _inventory_field(page: Page):
    return page.get_by_role("textbox", name=re.compile(r"Number of system inventory", re.I))

def _can_see_inventory_number_field(page: Page) -> bool:
    try:
        fld = _inventory_field(page)
        return fld.count() > 0 and fld.first.is_visible()
    except Exception:
        return False

def _can_see_single_record_button(page: Page) -> bool:
    try:
        btn = page.locator("div").filter(has_text=re.compile(r"^Single Record Entry$"))
        return btn.count() > 0 and btn.first.is_visible()
    except Exception:
        return False

def _state(page: Page) -> str:
    if _can_see_inventory_number_field(page):
        return "INVENTORY"
    if _can_see_single_record_button(page):
        return "INTERMEDIATE"
    return "UNKNOWN"

def _go_to_inventory_screen(page: Page):
    # Já está? nada a fazer. (Se houver telas intermediárias específicas, adicionar lógica aqui.)
    return

def _enter_inventory_number(page: Page, inv: str, max_attempts: int = 3):
    for attempt in range(1, max_attempts + 1):
        try:
            fld = _inventory_field(page)
            fld.first.wait_for(state="visible", timeout=4000)
            fld.first.click()
            try:
                fld.first.press("Control+A")
            except Exception:
                pass
            fld.first.fill(inv)
            log.info(f"[STEP] Inventory record -> '{inv}' (tentativa {attempt})")
            fld.first.press("Enter")
            time.sleep(0.9)
            if _can_see_single_record_button(page):
                return
        except Exception:
            log.debug("Tentativa falhou ao digitar/entrar inventário.")
        time.sleep(0.5)
    if not _can_see_single_record_button(page):
        raise ElementNotFound(f"Não chegou na tela intermediária para inventário '{inv}'.")

def _ensure_warehouse(page: Page):
    """
    Preenche Warehouse Number se campo existir e valor diferente.
    """
    try:
        fld = page.get_by_role("textbox", name=re.compile(r"Warehouse Number", re.I))
        if fld.count() > 0 and fld.first.is_visible():
            current = ""
            try:
                current = fld.first.input_value().strip()
            except Exception:
                pass
            if current.upper() != WAREHOUSE_VALUE:
                fld.first.click()
                try:
                    fld.first.press("Control+A")
                except Exception:
                    pass
                fld.first.fill(WAREHOUSE_VALUE)
                log.info(f"Warehouse definido: {WAREHOUSE_VALUE}")
    except Exception as e:
        log.debug(f"Warehouse não ajustado: {e}")

def _open_single_record_entry_after_inventory(page: Page, inv: str):
    _go_to_inventory_screen(page)
    _ensure_warehouse(page)
    _enter_inventory_number(page, inv)
    btn = page.locator("div").filter(has_text=re.compile(r"^Single Record Entry$"))
    btn.first.wait_for(state="visible", timeout=settings.DEFAULT_TIMEOUT)
    btn.first.click()
    log.info("Botão 'Single Record Entry' clicado.")
    _pause("after open SRE")
    # Aguarda campo Storage Bin (confirma entrada SRE)
    def _ready():
        try:
            loc = page.get_by_role("textbox", name="Storage Bin")
            return loc.count() > 0 and loc.first.is_visible()
        except Exception:
            return False
    try:
        wait_for(_ready, timeout_ms=settings.DEFAULT_TIMEOUT, action_desc="Campo 'Storage Bin' visível (SRE)")
        log.info("OK: Campo 'Storage Bin' visível (SRE)")
    except Exception:
        log.warning("Campo 'Storage Bin' não confirmado dentro do timeout.")