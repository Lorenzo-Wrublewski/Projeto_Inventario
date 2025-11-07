# config.py
# lib/config.py
# filepath: c:\Users\WRL1PO\Documents\Projeto_Inventario\lib\config.py
import os
import json
from dataclasses import dataclass
from typing import Any, Dict

CONFIG_JSON_FILENAME = "config.json"

def _load_json_config() -> Dict[str, Any]:
    path = os.path.join(os.getcwd(), CONFIG_JSON_FILENAME)
    if not os.path.isfile(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

_json_cfg = _load_json_config()

def _jget(path: str, default: Any):
    if not _json_cfg:
        return default
    ref = _json_cfg
    for part in path.split("."):
        if isinstance(ref, dict) and part in ref:
            ref = ref[part]
        else:
            return default
    return ref

@dataclass(frozen=True)
class Settings:
    BASE_URL: str = os.getenv(
        "SAP_BASE_URL",
        _jget("url", "http://rb3qraa0.server.bosch.com:8001/sap/bc/gui/sap/its/webgui/#")
    )
    HEADLESS: bool = (
        os.getenv("HEADLESS",
                  str(_jget("browser.headless", False))).lower()
        in ("1", "true", "yes")
    )
    SINGLE_RECORD_INTERVAL_S: float = float(
        os.getenv("SINGLE_RECORD_INTERVAL_S",
                  str(_jget("playback.single_record_interval_s", 1.0)))
    )
    SLOW_MO: int = int(
        os.getenv("SLOW_MO",
                  str(_jget("browser.slow_mo_ms", 0)))
    )
    DEFAULT_TIMEOUT: int = int(
        os.getenv("DEFAULT_TIMEOUT_MS",
                  str(_jget("timeouts.appear_ms", 15000)))
    )
    WAIT_POLL_INTERVAL: float = float(
        os.getenv("WAIT_POLL_INTERVAL",
                  str(_jget("timeouts.poll_interval_s", 0.25)))
    )
    PAGE_IDLE_TIMEOUT: int = int(
        os.getenv("PAGE_IDLE_TIMEOUT_MS",
                  str(int(_jget("timeouts.loading_timeout_s", 20) * 1000)))
    )
    RETRIES_ACTION: int = int(
        os.getenv("ACTION_RETRIES", "2")
    )
    SCREENSHOT_ON_ERROR: bool = (
        os.getenv("SHOT_ON_ERROR",
                  str(_jget("screenshots.on_error", True))).lower()
        in ("1", "true", "yes")
    )
    ACTION_DELAY_MS: int = int(
        os.getenv("ACTION_DELAY_MS",
                  str(_jget("playback.action_delay_ms", 0)))
    )
    VERBOSE_STEPS: bool = (
        os.getenv("VERBOSE_STEPS",
                  str(_jget("playback.verbose_steps", False))).lower()
        in ("1", "true", "yes")
    )
    FINAL_PAUSE_S: float = float(
        os.getenv("FINAL_PAUSE_S",
                  str(_jget("playback.final_pause_s", 0)))
    )
    REQUIRE_KEYPRESS_END: bool = (
        os.getenv("REQUIRE_KEYPRESS_END",
                  str(_jget("playback.require_keypress_end", False))).lower()
        in ("1", "true", "yes")
    )
    ZERO_STOCK_MODE: str = os.getenv(
        "ZERO_STOCK_MODE",
        _jget("playback.zero_stock_mode", "mark")
    ).lower()  # valores suportados: 'mark' ou 'skip'

settings = Settings()

def debug_print():
    for f in Settings.__dataclass_fields__:
        print(f"{f} = {getattr(settings, f)}")