# sap_actions.py
# c:\Users\WRL1PO\Documents\Projeto_Inventario\@Parte 1\lib\sap_actions.py
# filepath: c:\Users\WRL1PO\Documents\Projeto_Inventario\@Parte 1\lib\sap_actions.py
import re
import time
import os
from datetime import datetime
from playwright.sync_api import Page
from .logger import log
from .waits import safe_click, safe_fill, wait_until_any, wait_seconds, wait_for_interface_stable
from .config import Config

def narrar(msg: str):
    # Padroniza a “narração” das etapas
    log(f"[NARRAÇÃO] {msg}")

class SapSession:
    def __init__(self, page: Page, config: Config):
        self.page = page
        self.cfg = config

    def delay(self, etapa: str = ""):
        if self.cfg.action_delay > 0:
            narrar(f"Aguardando (delay configurado) - {etapa or 'Pausa'}")
            wait_seconds(self.cfg.action_delay, etapa or "Delay global")

    def tx_field(self):
        return self.page.get_by_role("textbox", name="Enter transaction code")

    def storage_field(self):
        return self.page.get_by_role("textbox", name="Storage Type")

    # ---------- NOVOS MÉTODOS DE SUPORTE (ROBUSTEZ) ----------

    def _is_storage_enabled(self) -> bool:
        try:
            field = self.storage_field()
            if not field.is_visible():
                return False
            disabled = bool(field.get_attribute("disabled") or field.get_attribute("aria-disabled"))
            readonly = bool(field.get_attribute("readonly"))
            return not disabled and not readonly
        except Exception:
            return False

    def _ensure_screen_ready(self, reason: str = ""):
        wait_time = getattr(self.cfg, "f8_ready_wait_seconds", 1.2)
        narrar(f"Verificando se a tela está pronta ({reason}) por até {wait_time:.1f}s")
        end = time.time() + wait_time
        while time.time() < end:
            if self._is_storage_enabled():
                narrar("Tela pronta para F8")
                return
            time.sleep(0.2)
        log(f"Prosseguindo sem confirmação total de readiness ({reason}).", level="WARN")

    def _refocus_before_f8(self):
        narrar("Garantindo foco antes de F8")
        try:
            refocus_selector = getattr(self.cfg, "refocus_selector", "")
            if refocus_selector:
                loc = self.page.locator(refocus_selector)
                if loc.count() > 0:
                    loc.first.click(timeout=1000)
                    narrar("Refoco via seletor customizado")
                    return
            field = self.storage_field()
            if field.is_visible():
                try:
                    field.click(timeout=500)
                    narrar("Refoco no campo Storage Type")
                except Exception:
                    pass
        except Exception:
            pass

    def _f8_effect_detected(self) -> bool:
        try:
            transfer_locator = self.page.get_by_role("cell", name="Transfer active", exact=True)
            if transfer_locator.is_visible():
                narrar("Detectado 'Transfer active' após F8")
                return True
        except Exception:
            pass
        try:
            activate_locator = self.page.locator("div").filter(has_text=re.compile(r"^Activate$"))
            if activate_locator.is_visible():
                narrar("Detectado botão 'Activate' após F8")
                return True
        except Exception:
            pass
        try:
            if not self.storage_field().is_visible():
                narrar("Campo Storage desapareceu - mudança de tela presumida")
                return True
        except Exception:
            return True
        return False

    def _revalidate_storage_if_needed(self, storage_code: str):
        if not getattr(self.cfg, "revalidate_storage_each_retry", True):
            return
        narrar(f"Revalidando Storage Type '{storage_code}' (Enter)")
        try:
            field = self.storage_field()
            if field.is_visible():
                field.click()
                field.press("Enter")
                wait_seconds(0.5, "Revalidação Storage (Enter)")
        except Exception:
            pass

    # ---------- FIM NOVOS MÉTODOS DE SUPORTE ----------

    def wait_transaction_field_ready(self):
        narrar("Aguardando campo de transação ficar pronto")
        end = time.time() + self.cfg.wait_tx_ready_timeout
        while time.time() < end:
            try:
                tx = self.tx_field()
                tx.wait_for(state="visible", timeout=2000)
                try:
                    tx.click(timeout=1000)
                except Exception:
                    pass
                disabled = bool(tx.get_attribute("disabled") or tx.get_attribute("aria-disabled"))
                readonly = bool(tx.get_attribute("readonly"))
                if not disabled and not readonly:
                    narrar("Campo de transação habilitado")
                    return
            except Exception:
                pass
            wait_seconds(1, "Aguardando campo transação habilitar")
        log("Timeout aguardando campo de transação.", level="WARN")

    def open_transaction(self, code: str):
        narrar(f"Iniciando abertura da transação '{code}'")
        self.wait_transaction_field_ready()
        narrar(f"Digitando transação '{code}'")
        safe_fill(self.tx_field(), code, description="Campo transação")
        self.delay("Antes Enter transação")
        narrar("Pressionando Enter para abrir transação")
        self.tx_field().press("Enter")
        self.delay("Após Enter transação")
        if code.upper() == "LX15":
            narrar("Aguardando campo Storage Type da LX15")
            try:
                self.storage_field().wait_for(timeout=20000)
                narrar("Tela LX15 carregada (Storage Type visível)")
            except Exception:
                log("Campo Storage Type não apareceu após abrir LX15.", level="WARN")

    def press_f8(self):
        attempts = getattr(self.cfg, "f8_multi_press_attempts", 2)
        narrar(f"Preparando para enviar F8 (até {attempts} disparos)")
        self._ensure_screen_ready("press_f8")
        self._refocus_before_f8()

        for i in range(1, attempts + 1):
            narrar(f"Enviando F8 (tentativa {i}/{attempts})")
            try:
                self.page.keyboard.press("F8")
            except Exception:
                try:
                    self.tx_field().press("F8")
                except Exception:
                    pass
            self.delay(f"Após F8 ({i})")
            time.sleep(0.4)
            if self._f8_effect_detected():
                narrar("Efeito de F8 detectado (parando tentativas iniciais)")
                return
        log("F8 aparentemente não produziu efeito imediato (continuará lógica de retry).", level="WARN")

    def choose_variant(self, variant_name: str):
        narrar(f"Abrindo lista de variantes para selecionar '{variant_name}'")
        variant_btn = self.page.locator("div").filter(has_text=re.compile(r"^Get Variant\.\.\.$"))
        safe_click(variant_btn, "Botão Get Variant")
        self.delay("Após abrir Get Variant")
        narrar(f"Selecionando linha da variante '{variant_name}'")
        row = self.page.get_by_role("row", name=variant_name, exact=True)
        safe_click(row.locator("div").nth(1), f"Linha variante {variant_name}")
        self.delay("Após selecionar linha variante")
        narrar("Confirmando variante (Choose)")
        choose_btn = self.page.get_by_title("Choose (F2)")
        safe_click(choose_btn, "Confirmar variante")
        self.delay("Após confirmar variante")

    def set_storage_type(self, storage_code: str):
        narrar(f"Preparando para digitar Storage Type '{storage_code}'")
        field = self.storage_field()
        field.wait_for(timeout=15000)
        try:
            narrar("Limpando campo Storage Type (Ctrl+A / Delete)")
            field.click()
            field.press("Control+A")
            field.press("Delete")
        except Exception:
            pass
        narrar(f"Digitando Storage Type '{storage_code}'")
        safe_fill(field, storage_code, description="Storage Type")
        self.delay("Antes Enter Storage Type")
        narrar("Confirmando Storage Type com Enter")
        field.press("Enter")
        self.delay("Após Enter Storage Type")
        narrar("Aguardando processamento do Storage Type")
        wait_seconds(1.5, "Processando Storage Type")

    def detect_transfer_or_activate_full(self):
        narrar("Detecção FULL pós F8 (Transfer active / Activate)")
        transfer_locator = self.page.get_by_role("cell", name="Transfer active", exact=True)
        activate_locator = self.page.locator("div").filter(has_text=re.compile(r"^Activate$"))
        try:
            desc, _ = wait_until_any(
                self.page,
                [
                    ("TRANSFER_ACTIVE", transfer_locator),
                    ("ACTIVATE_BUTTON", activate_locator),
                ],
                timeout=self.cfg.wait_after_f8_seconds
            )
            narrar(f"Detecção FULL encontrou: {desc}")
            return desc
        except Exception:
            log("Nenhuma condição encontrada (detecção full).", level="WARN")
            return None

    def detect_transfer_or_activate_quick(self):
        transfer_locator = self.page.get_by_role("cell", name="Transfer active", exact=True)
        time.sleep(2)
        activate_locator = self.page.locator("div").filter(has_text=re.compile(r"^Activate$"))
        end = time.time() + self.cfg.quick_detection_timeout
        narrar("Detecção QUICK pós F8 iniciada")
        while time.time() < end:
            try:
                if transfer_locator.is_visible():
                    narrar("QUICK detectou 'Transfer active'")
                    return "TRANSFER_ACTIVE"
            except Exception:
                pass
            try:
                if activate_locator.is_visible():
                    narrar("QUICK detectou 'Activate'")
                    return "ACTIVATE_BUTTON"
            except Exception:
                pass
            if not self._still_on_lx15_selection():
                narrar("Saiu da tela de seleção durante detecção QUICK")
                break
            time.sleep(self.cfg.quick_detection_poll)
        return None

    def click_activate_then_exit(self):
        narrar("Tentando clicar em Activate")
        time.sleep(2)
        activate_btn = self.page.locator("div").filter(has_text=re.compile(r"^Activate$"))
        if activate_btn.is_visible():
            safe_click(activate_btn, "Activate")
        else:
            log("Botão Activate não visível.", level="WARN")
        self.delay("Após Activate")
        narrar("Aguardando estabilização pós Activate")
        wait_seconds(1, "Estabilização pós Activate")
        self.exit_to_home()

    def exit_to_home(self):
        narrar("Iniciando sequência de Exit para voltar ao início")
        time.sleep(2)
        exit_btn = self.page.locator("div").filter(has_text=re.compile(r"^Exit$"))
        for i in range(3):
            if exit_btn.is_visible():
                narrar(f"Clicando Exit ({i+1})")
                safe_click(exit_btn, f"Exit ({i+1})")
                self.delay("Após Exit")
                wait_seconds(1, "Estabilização Exit")
            else:
                break

    def run_sm35_background_process(self):
        narrar("Abrindo transação SM35 para processamento em background")
        self.open_transaction("SM35")
        time.sleep(2)
        wait_seconds(2, "Carregando lista de batch inputs")
        narrar("Selecionando primeiro batch")
        safe_click(self.page.locator(".urST5SCMetricInner").first, "Primeiro batch")
        self.delay("Após selecionar batch")
        narrar("Clicando em Process")
        time.sleep(2)
        safe_click(self.page.locator("div").filter(has_text=re.compile(r"^Process$")), "Botão Process")
        self.delay("Após botão Process")
        narrar("Selecionando modo Background")
        time.sleep(2)
        safe_click(self.page.get_by_text("Background", exact=True), "Opção Background")
        self.delay("Após Background")
        narrar("Confirmando Process interno")
        time.sleep(2)
        inner_process = self.page.locator("#SAPMSBDC_CC300_1-tbcontainer div").filter(has_text=re.compile(r"^Process$"))
        safe_click(inner_process, "Process (Dentro do container)")
        self.delay("Após Process interno")
        time.sleep(2)
        exit_btn = self.page.locator("div").filter(has_text=re.compile(r"^Exit$"))
        if exit_btn.is_visible():
            narrar("Saindo da SM35 (Exit Final)")
            safe_click(exit_btn, "Exit Final")
            self.delay("Após Exit Final")
        narrar("Processo SM35 finalizado")

    def _save_transfer_active_screenshot(self, storage_code: str):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"transfer_active_{storage_code}_{ts}.png"
        path = os.path.join(self.cfg.error_screenshot_dir, filename)
        try:
            self.page.screenshot(path=path)
            log(f"Screenshot de 'Transfer active' salvo: {path}", level="WARN")
        except Exception as e:
            log(f"Falha ao salvar screenshot de erro ({e})", level="WARN")

    def _still_on_lx15_selection(self) -> bool:
        try:
            return self.storage_field().is_visible()
        except Exception:
            return False

    def _retry_f8_until_results(self, storage_code: str):
        attempt = 1
        max_attempts = self.cfg.f8_retry_attempts
        while attempt <= max_attempts:
            narrar(f"Verificando resultado pós F8 (tentativa lógica {attempt}/{max_attempts})")
            self._ensure_screen_ready(f"retry_f8_attempt_{attempt}")
            wait_seconds(self.cfg.wait_post_f8_small, "Pausa pós F8 (curta)")
            resultado = self.detect_transfer_or_activate_quick()
            if resultado in ("TRANSFER_ACTIVE", "ACTIVATE_BUTTON"):
                narrar(f"Resultado detectado: {resultado}")
                return resultado

            if not self._still_on_lx15_selection():
                narrar("Saiu da tela de seleção sem indicadores detectados")
                return None

            narrar("Ainda na tela de seleção sem indicadores")
            if attempt < max_attempts:
                self._revalidate_storage_if_needed(storage_code)
                self._refocus_before_f8()
                narrar("Reenviando F8 (ciclo de retry)")
                self.press_f8()
                wait_seconds(self.cfg.f8_retry_interval, "Intervalo entre tentativas F8")

            attempt += 1

        if self.cfg.final_full_detection:
            narrar("Executando detecção FULL final")
            return self.detect_transfer_or_activate_full()
        narrar("Encerrando tentativas de F8 sem resultado")
        return None

    def process_storage(self, storage_code: str):
        log(f"===== INÍCIO STORAGE {storage_code} =====")
        narrar(f"Iniciando processamento do Storage '{storage_code}'")
        try:
            narrar("Abrindo transação LX15")
            self.open_transaction("LX15")
            wait_seconds(2, "Carregando LX15")
            narrar(f"Selecionando variante '{self.cfg.variant_name}'")
            self.choose_variant(self.cfg.variant_name)
            wait_seconds(1, "Pausa pós variante")
            narrar(f"Configurando Storage Type '{storage_code}'")
            self.set_storage_type(storage_code)
            wait_seconds(1.0, "Estabilização antes de F8")
            narrar("Pressionando F8 para prosseguir")
            self.press_f8()

            narrar("Registrando screenshot pós F8")
            try:
                self.page.screenshot(path=self.cfg.screenshot_after_f8)
            except Exception:
                pass

            narrar("Iniciando verificação de resultados pós F8")
            resultado = self._retry_f8_until_results(storage_code)

            if resultado == "TRANSFER_ACTIVE":
                narrar(f"Storage '{storage_code}' já está em Transfer active (abortando este item)")
                self._save_transfer_active_screenshot(storage_code)
                self.exit_to_home()
                return "TRANSFER_ACTIVE"
            elif resultado == "ACTIVATE_BUTTON":
                narrar("Botão Activate disponível - ativando")
                self.click_activate_then_exit()
            else:
                narrar("Nenhum indicador detectado - efetuando saída")
                self.exit_to_home()

            narrar("Executando processamento em background (SM35)")
            self.run_sm35_background_process()
            narrar(f"Storage '{storage_code}' finalizado com sucesso")
            return "OK"
        except Exception as e:
            log(f"Erro inesperado no storage {storage_code}: {e}", level="WARN")
            narrar("Erro inesperado - capturando screenshot e retornando ERROR")
            try:
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                fname = f"exception_{storage_code}_{ts}.png"
                self.page.screenshot(path=os.path.join(self.cfg.error_screenshot_dir, fname))
            except Exception:
                pass
            self.exit_to_home()
            return "ERROR"
        finally:
            log(f"===== FIM STORAGE {storage_code} =====")
            narrar(f"Encerrado ciclo do Storage '{storage_code}'")