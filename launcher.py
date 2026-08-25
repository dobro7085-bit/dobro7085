from __future__ import annotations

import os
import socket
import sys
import threading
import time
import webbrowser
from pathlib import Path


APP_NAME = "AGM 감액량 분석"


def runtime_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def bundled_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent


def find_free_port(start_port: int = 8501, max_tries: int = 50) -> int:
    for port in range(start_port, start_port + max_tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind(("127.0.0.1", port))
            except OSError:
                continue
            return port
    raise RuntimeError("사용 가능한 포트를 찾지 못했습니다.")


def open_browser_later(url: str) -> None:
    time.sleep(2.5)
    webbrowser.open(url)


def ensure_external_standard_file() -> None:
    external_standard = runtime_dir() / "app" / "data" / "standard.xlsx"
    bundled_standard = bundled_dir() / "app" / "data" / "standard.xlsx"

    if external_standard.exists() or not bundled_standard.exists():
        return

    external_standard.parent.mkdir(parents=True, exist_ok=True)
    external_standard.write_bytes(bundled_standard.read_bytes())


def main() -> None:
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    os.environ.setdefault("STREAMLIT_SERVER_HEADLESS", "true")
    os.environ.setdefault("STREAMLIT_SERVER_FILE_WATCHER_TYPE", "none")
    os.environ.setdefault("STREAMLIT_SERVER_RUN_ON_SAVE", "false")
    os.environ.setdefault("STREAMLIT_SERVER_ENABLE_WEBSOCKET_COMPRESSION", "false")
    os.environ.setdefault("STREAMLIT_SERVER_WEBSOCKET_PING_INTERVAL", "10")
    os.environ.setdefault("STREAMLIT_SERVER_DISCONNECTED_SESSION_TTL", "3600")

    ensure_external_standard_file()

    app_path = bundled_dir() / "app.py"
    if not app_path.exists():
        app_path = runtime_dir() / "app.py"
    if not app_path.exists():
        raise FileNotFoundError("app.py 파일을 찾지 못했습니다.")

    port = find_free_port()
    url = f"http://localhost:{port}"
    threading.Thread(target=open_browser_later, args=(url,), daemon=True).start()

    from streamlit.web import cli as stcli

    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        "--server.address",
        "localhost",
        "--server.port",
        str(port),
        "--server.headless",
        "true",
        "--server.fileWatcherType",
        "none",
        "--server.runOnSave",
        "false",
        "--server.enableWebsocketCompression",
        "false",
        "--server.websocketPingInterval",
        "10",
        "--server.disconnectedSessionTTL",
        "3600",
        "--global.developmentMode",
        "false",
        "--browser.gatherUsageStats",
        "false",
    ]

    print(f"{APP_NAME} 실행 중입니다.")
    print("이 창은 프로그램 서버입니다. 사용 중에는 닫지 마세요.")
    print(f"브라우저가 자동으로 열리지 않으면 아래 주소로 접속하세요: {url}")
    stcli.main()


if __name__ == "__main__":
    main()
