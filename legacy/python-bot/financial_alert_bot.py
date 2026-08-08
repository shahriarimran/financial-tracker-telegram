from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import time
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import pythoncom
import requests
import win32com.client
from dotenv import load_dotenv


LOGGER = logging.getLogger("financial-alert-bot")

# Cells already present in the supplied workbook.
ASSET_CELLS = {
    "Gold": {
        "status": "N16",
        "buy": "H16",
        "sell": "I16",
        "spread": "K16",
        "daily_change": "K24",
        "unit": "TRY/g",
    },
    "Silver": {
        "status": "N18",
        "buy": "H18",
        "sell": "I18",
        "spread": "K18",
        "daily_change": "K25",
        "unit": "TRY/g",
    },
    "USD": {
        "status": "N20",
        "buy": "H20",
        "sell": "I20",
        "spread": "K20",
        "daily_change": "K23",
        "unit": "TRY/USD",
    },
}


@dataclass(frozen=True)
class AssetSnapshot:
    status: str
    buy: float | None
    sell: float | None
    spread: float | None
    daily_change: float | None
    unit: str


def _as_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _normalize_status(value: Any) -> str:
    if value is None:
        return "UNKNOWN"
    text = str(value).strip().upper()
    return text or "UNKNOWN"


class ExcelSignalReader:
    """Open the macro-enabled workbook in an isolated Excel instance."""

    def __init__(self, workbook_path: Path, sheet_name: str = "Financial Tracker") -> None:
        self.workbook_path = workbook_path
        self.sheet_name = sheet_name
        self.excel = None
        self.workbook = None

    def __enter__(self) -> "ExcelSignalReader":
        pythoncom.CoInitialize()
        try:
            self.excel = win32com.client.DispatchEx("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False

            # Open read-only. Power Query can refresh in memory; the source workbook
            # is not modified by the alert bot.
            self.workbook = self.excel.Workbooks.Open(
                str(self.workbook_path),
                UpdateLinks=3,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
            )
            return self
        except Exception:
            self._close()
            raise

    def __exit__(self, exc_type, exc, tb) -> None:
        self._close()

    def _close(self) -> None:
        try:
            if self.workbook is not None:
                self.workbook.Close(SaveChanges=False)
        except Exception:
            pass

        try:
            if self.excel is not None:
                self.excel.Quit()
        except Exception:
            pass

        self.workbook = None
        self.excel = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    def refresh_and_read(
        self,
        refresh_timeout_seconds: int = 90,
    ) -> tuple[dict[str, AssetSnapshot], Any]:

        if self.workbook is None or self.excel is None:
            raise RuntimeError("Excel workbook is not open.")

        LOGGER.info("Refreshing Excel queries...")

        # Start Power Query / workbook refresh.
        self.workbook.RefreshAll()

        # For this workbook, manual Refresh All completes normally.
        # Avoid querying Connection.Refreshing through COM because some
        # Power Query connections can block Python indefinitely.
        refresh_wait = min(refresh_timeout_seconds, 20)

        LOGGER.info(
            "Waiting %d seconds for Power Query refresh...",
            refresh_wait,
        )

        time.sleep(refresh_wait)

        LOGGER.info("Recalculating workbook...")

        try:
            self.excel.CalculateFullRebuild()
        except Exception:
            self.excel.CalculateFull()

        time.sleep(2)

        sheet = self.workbook.Worksheets(self.sheet_name)

        snapshots: dict[str, AssetSnapshot] = {}

        for asset, cells in ASSET_CELLS.items():
            snapshots[asset] = AssetSnapshot(
                status=_normalize_status(
                    sheet.Range(cells["status"]).Value
                ),
                buy=_as_float(
                    sheet.Range(cells["buy"]).Value
                ),
                sell=_as_float(
                    sheet.Range(cells["sell"]).Value
                ),
                spread=_as_float(
                    sheet.Range(cells["spread"]).Value
                ),
                daily_change=_as_float(
                    sheet.Range(cells["daily_change"]).Value
                ),
                unit=str(cells["unit"]),
            )

        workbook_timestamp = sheet.Range("G1").Value

        LOGGER.info(
            "Signals read: %s",
            ", ".join(
                f"{asset}={snapshot.status}"
                for asset, snapshot in snapshots.items()
            ),
        )

        return snapshots, workbook_timestamp


def send_telegram_message(token: str, chat_id: str, text: str) -> None:
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    response = requests.post(
        url,
        data={
            "chat_id": chat_id,
            "text": text,
            "disable_web_page_preview": True,
        },
        timeout=20,
    )
    response.raise_for_status()
    payload = response.json()
    if not payload.get("ok", False):
        raise RuntimeError(f"Telegram API returned an error: {payload}")


def load_state(path: Path) -> dict[str, dict[str, Any]]:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        LOGGER.warning("Could not read state file; treating this as first run.")
        return {}


def save_state(path: Path, snapshots: dict[str, AssetSnapshot]) -> None:
    payload = {name: asdict(snapshot) for name, snapshot in snapshots.items()}
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    tmp.replace(path)


def fmt_number(value: float | None, digits: int = 4) -> str:
    if value is None:
        return "n/a"
    return f"{value:,.{digits}f}".rstrip("0").rstrip(".")


def fmt_pct(value: float | None) -> str:
    if value is None:
        return "n/a"
    return f"{value * 100:+.2f}%"


def fmt_workbook_time(value: Any) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    return str(value) if value not in (None, "") else "unknown"


def asset_line(asset: str, snap: AssetSnapshot) -> str:
    return (
        f"{asset}: {snap.status} | "
        f"buy {fmt_number(snap.buy)} | sell {fmt_number(snap.sell)} {snap.unit} | "
        f"spread {fmt_pct(snap.spread)} | daily {fmt_pct(snap.daily_change)}"
    )


def startup_message(snapshots: dict[str, AssetSnapshot], workbook_time: Any) -> str:
    lines = [
        "Financial alert bot started.",
        f"Workbook timestamp: {fmt_workbook_time(workbook_time)}",
        "",
    ]
    lines.extend(asset_line(asset, snap) for asset, snap in snapshots.items())
    return "\n".join(lines)


def change_message(
    changed_assets: list[tuple[str, str, AssetSnapshot]],
    workbook_time: Any,
) -> str:
    lines = [
        "Financial signal changed",
        f"Workbook timestamp: {fmt_workbook_time(workbook_time)}",
        "",
    ]
    for asset, previous_status, current in changed_assets:
        lines.append(f"{asset}: {previous_status} -> {current.status}")
        lines.append(
            f"  buy {fmt_number(current.buy)} | sell {fmt_number(current.sell)} {current.unit}"
        )
        lines.append(
            f"  spread {fmt_pct(current.spread)} | daily {fmt_pct(current.daily_change)}"
        )
    return "\n".join(lines)


def detect_changes(
    previous: dict[str, dict[str, Any]],
    current: dict[str, AssetSnapshot],
) -> list[tuple[str, str, AssetSnapshot]]:
    changed: list[tuple[str, str, AssetSnapshot]] = []
    for asset, snapshot in current.items():
        old_status = _normalize_status(previous.get(asset, {}).get("status"))
        if old_status != snapshot.status:
            changed.append((asset, old_status, snapshot))
    return changed


def run_once(
    workbook_path: Path,
    token: str,
    chat_id: str,
    state_path: Path,
    send_startup: bool,
    refresh_timeout_seconds: int,
) -> None:
    with ExcelSignalReader(workbook_path) as reader:
        snapshots, workbook_time = reader.refresh_and_read(refresh_timeout_seconds)

    previous = load_state(state_path)

    if not previous:
        LOGGER.info("First run: storing baseline state.")
        save_state(state_path, snapshots)
        if send_startup:
            send_telegram_message(token, chat_id, startup_message(snapshots, workbook_time))
        return

    changed = detect_changes(previous, snapshots)
    if changed:
        message = change_message(changed, workbook_time)
        send_telegram_message(token, chat_id, message)
        LOGGER.info("Sent Telegram alert for %d changed asset(s).", len(changed))
    else:
        LOGGER.info(
            "No signal change: %s",
            ", ".join(f"{a}={s.status}" for a, s in snapshots.items()),
        )

    # Save after Telegram succeeds so a transient Telegram failure does not silently
    # consume an alert.
    save_state(state_path, snapshots)


def get_required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def main() -> int:
    load_dotenv()

    parser = argparse.ArgumentParser(
        description="Refresh an Excel financial tracker and alert Telegram on BUY/WAIT state changes."
    )
    parser.add_argument("--once", action="store_true", help="Check once, then exit.")
    parser.add_argument(
        "--test-telegram",
        action="store_true",
        help="Send a Telegram test message, then exit.",
    )
    args = parser.parse_args()

    token = get_required_env("TELEGRAM_BOT_TOKEN")
    chat_id = get_required_env("TELEGRAM_CHAT_ID")

    if args.test_telegram:
        send_telegram_message(token, chat_id, "Telegram financial alert bot test: OK")
        print("Telegram test message sent.")
        return 0

    workbook_path = Path(get_required_env("WORKBOOK_PATH")).expanduser()
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    interval = int(os.getenv("CHECK_INTERVAL_SECONDS", "300"))
    refresh_timeout = int(os.getenv("REFRESH_TIMEOUT_SECONDS", "90"))
    send_startup = os.getenv("SEND_STARTUP_MESSAGE", "false").strip().lower() in {
        "1",
        "true",
        "yes",
        "y",
    }

    state_path = Path(
        os.getenv("STATE_FILE", str(Path(__file__).with_name("signal_state.json")))
    ).expanduser()

    log_path = Path(
        os.getenv("LOG_FILE", str(Path(__file__).with_name("financial_alert_bot.log")))
    ).expanduser()
    log_path.parent.mkdir(parents=True, exist_ok=True)

    logging.basicConfig(
        level=os.getenv("LOG_LEVEL", "INFO").upper(),
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_path, encoding="utf-8"),
        ],
    )

    if args.once:
        run_once(
            workbook_path,
            token,
            chat_id,
            state_path,
            send_startup,
            refresh_timeout,
        )
        return 0

    LOGGER.info("Bot running. Poll interval: %d seconds.", interval)

    last_error_text = None
    while True:
        try:
            run_once(
                workbook_path,
                token,
                chat_id,
                state_path,
                send_startup,
                refresh_timeout,
            )
            last_error_text = None
        except KeyboardInterrupt:
            LOGGER.info("Stopped by user.")
            return 0
        except Exception as exc:
            LOGGER.exception("Check failed.")
            error_text = f"{type(exc).__name__}: {exc}"

            # Avoid sending the same error every polling cycle.
            if error_text != last_error_text:
                try:
                    send_telegram_message(
                        token,
                        chat_id,
                        "Financial alert bot error:\n" + error_text,
                    )
                except Exception:
                    LOGGER.exception("Could not send Telegram error notification.")
                last_error_text = error_text

        time.sleep(max(interval, 30))


if __name__ == "__main__":
    raise SystemExit(main())
