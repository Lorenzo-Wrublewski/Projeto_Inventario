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
            df = pd.read_excel(path_str)
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

def _click_cancel_once(page: Page, confirm_yes: bool = False, wait_visible: float = 5.0):
    cancel_locator = page.locator("div").filter(has_text=re.compile(r"^Cancel$"))
    if cancel_locator.count() == 0:
        btn = page.get_by_role("button", name=re.compile(r"^Cancel$", re.I))
        if btn.count() > 0:
            cancel_locator = btn
    end = time.time() + wait_visible
    while time.time() < end:
        try:
            if cancel_locator.count() > 0 and cancel_locator.first.is_visible():
                cancel_locator.first.click()
                log.info("Click Cancel")
                time.sleep(0.35)
                break
        except Exception:
            pass
        time.sleep(0.2)
    else:
        log.warning("Não foi possível clicar em Cancel (não visível).")
    if confirm_yes:
        try:
            yes_btn = page.get_by_title("Yes")
            if yes_btn.is_visible():
                yes_btn.click()
                log.info("Click Yes (confirmação)")
                time.sleep(0.4)
        except Exception:
            pass

def _can_see_inventory_number_field(page: Page) -> bool:
    try:
        fld = page.get_by_role("textbox", name=re.compile(r"Number of system inventory", re.I))
        return fld.count() > 0 and fld.first.is_visible()
    except Exception:
        return False

def _can_see_single_record_button(page: Page) -> bool:
    try:
        time.sleep(1)
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

def _final_cancel_to_inventory(page: Page):
    st = _state(page)
    log.info(f"Estado antes do Cancel final: {st}")
    if st == "INVENTORY":
        log.info("Já em INVENTORY.")
    _click_cancel_once(page, confirm_yes=True)
    time.sleep(0.6)
    st2 = _state(page)
    log.info(f"Estado após Cancel final: {st2}")
    if st2 != "INVENTORY":
        log.warning("Tentando segundo Cancel.")
        _click_cancel_once(page, confirm_yes=True)
        time.sleep(0.6)
        st3 = _state(page)
        log.info(f"Estado após segunda tentativa: {st3}")
        if st3 != "INVENTORY":
            log.warning("Não retornou INVENTORY.")

ACTION_INTERVAL_S = getattr(settings, "SINGLE_RECORD_INTERVAL_S", 0.4)
SHORT_SLEEP = getattr(settings, "SHORT_SLEEP_S", 0.25)

def _pause(tag: str = ""):
    time.sleep(ACTION_INTERVAL_S)

def _inventory_field(page: Page):
    return page.get_by_role("textbox", name=re.compile(r"Number of system inventory", re.I))

def _go_to_inventory_screen(page: Page):
    if _state(page) != "INVENTORY":
        _final_cancel_to_inventory(page)

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
        time.sleep(0.6)
    if not _can_see_single_record_button(page):
        raise ElementNotFound(f"Não chegou na tela intermediária para inventário '{inv}'.")

def _wait_until_storage_bin_field(page: Page, timeout_ms: int | None = None):
    """
    Aguarda campo 'Storage Bin' ficar visível, garantindo que a tela Single Record Entry carregou.
    """
    timeout_ms = timeout_ms or settings.DEFAULT_TIMEOUT
    def _ready():
        try:
            loc = page.get_by_role("textbox", name="Storage Bin")
            return loc.count() > 0 and loc.first.is_visible()
        except Exception:
            return False
    wait_for(_ready, timeout_ms=timeout_ms, action_desc="Campo 'Storage Bin' visível (SRE)")

def _wait_until_inventory_screen(page: Page, timeout_ms: int | None = None):
    """
    Aguarda retorno à tela de inventário (campo 'Number of system inventory' visível).
    """
    timeout_ms = timeout_ms or settings.DEFAULT_TIMEOUT
    def _inv():
        return _state(page) == "INVENTORY"
    wait_for(_inv, timeout_ms=timeout_ms, action_desc="Tela INVENTORY disponível")

def _open_single_record_entry_after_inventory(page: Page, inv: str):
    _go_to_inventory_screen(page)
    _enter_inventory_number(page, inv)
    btn = _single_record_button_locator(page)
    btn.first.wait_for(state="visible", timeout=settings.DEFAULT_TIMEOUT)
    btn.first.click()
    log.info("Botão 'Single Record Entry' clicado.")
    _pause("after open SRE")
    # NOVO: aguarda os campos da SRE realmente carregarem
    _wait_until_storage_bin_field(page)

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

def process_single_record_entries(page: Page, contagem_path: str, reference_report_path: Optional[str] = None):
    """
    Nova lógica:
    - Arquivo de contagem: uma linha por material (total contado).
    - Relatório referência: várias linhas (UD ou não).
    - Ajusta apenas linhas com UD; linhas sem UD lançadas sem alteração.
    """
    contagem_records = load_single_record_file(contagem_path)
    if not contagem_records:
        log.warning("Nenhum registro de contagem.")
        return

    reference_records = load_comparison_report(reference_report_path) if reference_report_path else []
    if not reference_records:
        log.warning("Relatório de referência vazio. Lançando contagem sem lógica UD.")
        # Fallback: cada linha de contagem vira um lançamento direto
        for idx, rec in enumerate(contagem_records, start=1):
            # Usa quantidade contada diretamente
            qty_val = rec.get("counted_quantity") or rec.get("quantity_alt") or "0"
            rec["quantity_alt"] = _format_quantity(_parse_number(qty_val))
            _process_single_record(page, rec, idx, len(contagem_records))
        return

    total_contagem = len(contagem_records)
    log.info(f"Processando {total_contagem} materiais (modo nova lógica UD).")

    launch_list: List[Dict[str, str]] = []

    for rec_index, cont in enumerate(contagem_records, start=1):
        material = cont.get("material_number", "").strip()
        plant = cont.get("plant", "").strip()
        storage_location = cont.get("storage_location", "").strip()
        storage_type = cont.get("storage_type", "").strip()
        storage_bin = cont.get("storage_bin", "").strip()
        inventory_record = cont.get("inventory_record", "").strip()

        counted_total = _parse_number(cont.get("counted_quantity") or cont.get("quantity_alt") or "0")
        log.info(f"[MAT {rec_index}/{total_contagem}] Material={material} Bin={storage_bin} Total contado={counted_total}")

        # Filtra linhas referência compatíveis (se campo do contagem estiver vazio, não filtra por ele)
        def _match(ref: Dict[str, str]) -> bool:
            if ref.get("material_number", "").strip() != material:
                return False
            if plant and ref.get("plant", "").strip() != plant:
                return False
            if storage_location and ref.get("storage_location", "").strip() != storage_location:
                return False
            if storage_type and ref.get("storage_type", "").strip() != storage_type:
                return False
            if storage_bin and ref.get("storage_bin", "").strip() != storage_bin:
                return False
            return True

        matching = [r for r in reference_records if _match(r)]
        if not matching:
            log.warning(f"[MAT {material}] Nenhuma linha referência correspondente. Lançando registro único.")
            cont["quantity_alt"] = _format_quantity(counted_total)
            launch_list.append(cont)
            continue

        ud_rows = [r for r in matching if r.get("ud")]
        non_ud_rows = [r for r in matching if not r.get("ud")]

        log.info(f"[MAT {material}] {len(ud_rows)} UDs / {len(non_ud_rows)} linhas sem UD.")

        # Ajuste UD
        if ud_rows:
            adjusted_values = _adjust_ud_quantities(counted_total, ud_rows)
            # Ordena novamente para associar
            ordered_ud = sorted(ud_rows, key=lambda r: _parse_ud_number(r.get("ud", "")))
            soma_final = sum(adjusted_values)
            log.info(f"[MAT {material}] Soma UD ajustada={soma_final} (esperado={counted_total}) DeltaFinal={counted_total - (sum(_parse_number(r.get('stock_total','')) for r in ud_rows))}")
            for i, r_ud in enumerate(ordered_ud):
                qty_final = adjusted_values[i]
                formatted = _format_quantity(qty_final)
                # Monta registro para lançamento
                launch_rec = {
                    "inventory_record": inventory_record,
                    "material_number": r_ud.get("material_number", ""),
                    "storage_bin": r_ud.get("storage_bin", ""),
                    "plant": r_ud.get("plant", plant),
                    "storage_location": r_ud.get("storage_location", storage_location),
                    "storage_type": r_ud.get("storage_type", storage_type),
                    "ud": r_ud.get("ud", ""),
                    "counted_quantity": formatted,
                    "quantity_alt": formatted
                }
                is_zero = qty_final == 0
                if is_zero and settings.ZERO_STOCK_MODE == "skip":
                    log.info(f"[MAT {material}] UD {r_ud.get('ud')} ignorada (zero, modo skip).")
                    continue
                launch_list.append(launch_rec)
        else:
            # Sem UDs: lança contagem total como único registro
            cont["quantity_alt"] = _format_quantity(counted_total)
            launch_list.append(cont)

        # Linhas sem UD (lançamento normal com stock_total)
        for r_nu in non_ud_rows:
            qty_orig = _parse_number(r_nu.get("stock_total", "0"))
            formatted = _format_quantity(qty_orig)
            launch_rec = {
                "inventory_record": inventory_record,
                "material_number": r_nu.get("material_number", material),
                "storage_bin": r_nu.get("storage_bin", storage_bin),
                "plant": r_nu.get("plant", plant),
                "storage_location": r_nu.get("storage_location", storage_location),
                "storage_type": r_nu.get("storage_type", storage_type),
                "ud": "",
                "counted_quantity": formatted,
                "quantity_alt": formatted
            }
            is_zero = qty_orig == 0
            if is_zero and settings.ZERO_STOCK_MODE == "skip":
                log.info(f"[MAT {material}] Linha sem UD ignorada (zero, modo skip).")
                continue
            launch_list.append(launch_rec)

    # Lançamento efetivo
    total_launch = len(launch_list)
    log.info(f"Iniciando lançamento de {total_launch} linhas (UD ajustadas + sem UD).")

    for idx, rec in enumerate(launch_list, start=1):
        _process_single_record(page, rec, idx, total_launch)

    log.info("Processo concluído (nova lógica UD).")

def _process_single_record(page: Page, rec: Dict[str, str], idx: int, total: int, seq_info: Optional[str] = None):
    inv = rec.get("inventory_record", "").strip()
    if not inv:
        log.error(f"[Registro {idx}] Sem 'inventory_record'. Pulando.")
        return
    tag = f"[Registro {idx}/{total}]{'[' + seq_info + ']' if seq_info else ''}"
    try:
        _open_single_record_entry_after_inventory(page, inv)
        log.info(f"{tag} {rec}")

        time.sleep(1)
        _fill_field(page, "Storage Bin", rec.get("storage_bin", ""))
        _pause()
        time.sleep(SHORT_SLEEP)
        _fill_field(page, "Material Number", rec.get("material_number", ""))
        _pause()
        time.sleep(SHORT_SLEEP)
        _fill_field(page, "Counted quantity in alternative unit of measure",
                    _format_quantity(_parse_number(rec.get("quantity_alt") or rec.get("counted_quantity") or "0")))
        _pause()
        time.sleep(SHORT_SLEEP)

        qty_zero = rec.get("quantity_alt", "").strip() in ["0", "0.0", "0,0"]
        if qty_zero:
            try:
                page.get_by_text("Zero stock").click()
                log.info(f"{tag} Caixa 'Zero stock' marcada.")
                time.sleep(1)
            except Exception:
                log.warning(f"{tag} Não localizou 'Zero stock'.")

        _fill_field(page, "Storage Location", rec.get("storage_location", ""))
        _pause()
        time.sleep(SHORT_SLEEP)
        _fill_field(page, "Plant", rec.get("plant", ""))
        _pause()
        time.sleep(SHORT_SLEEP)

        qty_field = page.get_by_role("textbox", name="Counted quantity in alternative unit of measure")
        time.sleep(1)
        qty_field.press("Enter")
        _pause("after first enter qty")
        qty_field.press("Enter")
        _pause("after second enter qty")

        _click_cancel_once(page, confirm_yes=False)
        _pause("after first cancel")
        _click_cancel_once(page, confirm_yes=True)
        _pause("after second cancel")

        # Em vez de só checar, aguarda de fato INVENTORY
        try:
            _wait_until_inventory_screen(page, timeout_ms=settings.DEFAULT_TIMEOUT)
        except Exception:
            log.warning(f"{tag} Inventário não confirmado por timeout. Tentando forçar Cancel extra.")
            _final_cancel_to_inventory(page)
            try:
                _wait_until_inventory_screen(page, timeout_ms=settings.DEFAULT_TIMEOUT)
            except Exception:
                log.error(f"{tag} Falha persistente ao voltar INVENTORY.")

        log.info(f"{tag} concluído.")
    except ElementNotFound as e:
        log.error(f"{tag} Falha elemento: {e}")
        _final_cancel_to_inventory(page)
    except Exception as e:
        log.error(f"{tag} Erro inesperado: {e}")
        _final_cancel_to_inventory(page)