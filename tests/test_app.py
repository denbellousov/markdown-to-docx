"""Приложение md2docx (md2docx_app.py): с аргументами — как md2docx_web.py; без них — сервер и окно,
выход из окна останавливает сервер и удаляет временные файлы; второй запуск открывает первый.
Окно Tk здесь подменено: его проверяют вручную и дымовой тест сборки (packaging/smoke.py).

    python3 -m pytest tests/test_app.py -q --basetemp=<свой каталог> -p no:cacheprovider
"""

from __future__ import annotations

import http.client
import os
import signal
import socket
import subprocess
import sys
import threading
import time
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import md2docx_app as app  # noqa: E402


def free_port():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def get(url):
    host, port = url.split("//")[1].strip("/").split(":")
    conn = http.client.HTTPConnection(host, int(port), timeout=10)
    try:
        conn.request("GET", "/")
        resp = conn.getresponse()
        return resp.status, resp.read().decode()
    finally:
        conn.close()


@pytest.fixture
def opened(monkeypatch, tmp_path):
    """Браузер не открывается — адреса копятся в списке; временные каталоги — в tmp_path."""
    urls = []
    monkeypatch.setattr(app.webbrowser, "open", lambda url: urls.append(url) or True)
    monkeypatch.setattr(app.web.tempfile, "tempdir", str(tmp_path))
    return urls


@pytest.fixture(autouse=True)
def signals():
    """app() ставит свои обработчики SIGTERM и SIGHUP — после теста возвращаются прежние."""
    saved = {s: signal.getsignal(s) for s in app.STOP_SIGNALS}
    yield
    for s, handler in saved.items():
        signal.signal(s, handler)


def test_window_gets_a_live_server_and_quit_cleans_up(opened, monkeypatch, tmp_path):
    seen = {}

    def window(url):
        seen.update(url=url, page=get(url), dirs=list(tmp_path.iterdir()))

    port = free_port()
    monkeypatch.setattr(app.web, "PORT", port)
    monkeypatch.setattr(app, "window", window)
    assert app.app() == 0
    assert seen["url"] == f"http://127.0.0.1:{port}/"
    status, page = seen["page"]
    assert status == 200 and "const APP = true;" in page, "страница не знает, что её запустило приложение"
    assert len(seen["dirs"]) == 1 and list(tmp_path.iterdir()) == [], "временный каталог не удалён"
    with pytest.raises(OSError):
        get(seen["url"])  # сервер остановлен


def test_window_failure_still_cleans_up(opened, monkeypatch, tmp_path):
    def window(url):
        raise RuntimeError("no display")

    monkeypatch.setattr(app.web, "PORT", free_port())
    monkeypatch.setattr(app, "window", window)
    with pytest.raises(RuntimeError):
        app.app()
    assert list(tmp_path.iterdir()) == []


def test_kill_stops_like_quit(opened, monkeypatch, tmp_path):
    """kill (SIGTERM) — как «Выйти»: сервер остановлен, временные файлы удалены."""
    def window(url):
        os.kill(os.getpid(), signal.SIGTERM)
        time.sleep(10)  # обработчик прерывает ожидание

    signal.signal(signal.SIGTERM, lambda *_: pytest.fail("SIGTERM не перехвачен: временные файлы остались бы"))
    monkeypatch.setattr(app.web, "PORT", free_port())
    monkeypatch.setattr(app, "window", window)
    with pytest.raises(SystemExit) as stop:
        app.app()
    assert stop.value.code == 0 and list(tmp_path.iterdir()) == []


def test_second_launch_opens_the_running_one(opened, monkeypatch, tmp_path):
    first = app.web.Server(free_port(), app=True)
    threading.Thread(target=first.serve_forever, daemon=True).start()
    try:
        monkeypatch.setattr(app.web, "PORT", first.port)
        monkeypatch.setattr(app, "window", lambda url: pytest.fail("второе окно"))
        assert app.app() == 0
        assert opened == [f"http://127.0.0.1:{first.port}/"]
        assert list(tmp_path.iterdir()) == [first.root], "второй запуск создал свой сервер"
    finally:
        first.shutdown()
        first.server_close()
        first.cleanup()


def test_port_held_by_another_program_gives_any_free_port(opened, monkeypatch):
    urls = []
    with socket.socket() as other:
        other.bind(("127.0.0.1", 0))
        other.listen()
        port = other.getsockname()[1]
        monkeypatch.setattr(app.web, "PORT", port)
        monkeypatch.setattr(app, "window", urls.append)
        assert app.app() == 0
    (url,) = urls
    assert url.startswith("http://127.0.0.1:") and not url.endswith(f":{port}/")
    assert opened == []


def test_finder_process_serial_number_is_ignored(monkeypatch):
    monkeypatch.setattr(app, "app", lambda: "window")
    assert app.main(["-psn_0_12345"]) == "window"


def test_files_dropped_on_the_app_open_the_window(monkeypatch):
    """Windows передаёт файл, брошенный на значок, аргументом: окно, а не молчаливая ошибка argparse."""
    monkeypatch.setattr(app, "app", lambda: "window")
    assert app.main([r"C:\Users\me\report.md"]) == "window"


def test_arguments_run_the_web_interface(tmp_path):
    port = free_port()
    env = dict(os.environ, TMPDIR=str(tmp_path), PYTHONUNBUFFERED="1")
    proc = subprocess.Popen([sys.executable, str(ROOT / "md2docx_app.py"), "--port", str(port), "--no-browser"],
                            env=env, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    timer = threading.Timer(60, proc.kill)
    timer.start()
    try:
        assert f"http://127.0.0.1:{port}/" in proc.stdout.readline()
        status, page = get(f"http://127.0.0.1:{port}/")
        assert status == 200 and "const APP = false;" in page  # из терминала — подсказки про Ctrl+C
        proc.send_signal(signal.SIGTERM)
        assert proc.wait(timeout=30) == 0
    finally:
        timer.cancel()
        if proc.poll() is None:
            proc.kill()
        proc.communicate()
    assert list(tmp_path.iterdir()) == []
    shown = subprocess.run([sys.executable, str(ROOT / "md2docx_app.py"), "--help"],
                           capture_output=True, text=True, timeout=60)
    assert shown.returncode == 0 and shown.stdout.startswith(app.tr("usage:") + " md2docx ")


def test_missing_output_streams_are_replaced():
    """Оконная сборка под Windows запускается без stdout и stderr: print и журнал сервера не падают."""
    code = ("import sys; sys.stdout = sys.stderr = None; sys.path.insert(0, %r); import md2docx_app; "
            "print('lost'); sys.__stdout__.write(repr((sys.stdout.name, sys.stderr.name)))" % str(ROOT))
    proc = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, timeout=60)
    assert proc.returncode == 0 and proc.stdout == repr((os.devnull, os.devnull)), proc.stderr
