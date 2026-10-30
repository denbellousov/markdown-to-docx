"""Приложение md2docx (md2docx_app.py): с аргументами — как md2docx_web.py; без них — сервер и своё окно,
закрытие окна, «Завершить» и kill останавливают сервер и удаляют временные файлы.
Окно pywebview здесь подменено: настоящее проверяют вручную и дымовой тест сборки (packaging/smoke.py).

    python3 -m pytest tests/test_app.py -q --basetemp=<свой каталог> -p no:cacheprovider
"""

from __future__ import annotations

import ctypes
import http.client
import os
import signal
import socket
import subprocess
import sys
import tempfile
import threading
import time
import types
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


def serving(url):
    try:
        return get(url)[0] == 200
    except OSError:
        return False


@pytest.fixture
def tmp_dirs(monkeypatch, tmp_path):
    """Временные каталоги сервера — в tmp_path."""
    monkeypatch.setattr(app.web.tempfile, "tempdir", str(tmp_path))
    return tmp_path


@pytest.fixture(autouse=True)
def signals():
    """on_stop_signals() ставит свои обработчики SIGTERM и SIGHUP и wakeup_fd — после теста возвращаются прежние."""
    saved = {s: signal.getsignal(s) for s in app.STOP_SIGNALS}
    wakeup = signal.set_wakeup_fd(-1)
    signal.set_wakeup_fd(wakeup)
    yield
    signal.set_wakeup_fd(wakeup)
    for s, handler in saved.items():
        signal.signal(s, handler)


class Event(list):
    def __iadd__(self, handler):
        self.append(handler)
        return self


@pytest.fixture
def webview(monkeypatch):
    """Подмена pywebview: start() сначала зовёт обработчики initialized (renderer — как у настоящего окна),
    затем, если окно показано, — обработчики closing, как при «Завершить» в меню, и запоминает,
    работает ли сервер и что осталось во временном каталоге: после «Завершить» finally в app() не будет."""
    seen = types.SimpleNamespace(renderer="cocoa", settings={}, windows=[], started=None, after_closing=None,
                                 on_stop=None)

    def create_window(*args, **kwargs):
        win = types.SimpleNamespace(args=args, kwargs=kwargs, destroy=lambda: None,
                                    events=types.SimpleNamespace(initialized=Event(), closing=Event()))
        seen.windows.append(win)
        return win

    def start(**kwargs):
        seen.started = kwargs
        (win,) = seen.windows
        if all(handler(seen.renderer) is not False for handler in win.events.initialized):
            for handler in win.events.closing:
                handler()
            seen.after_closing = (serving(win.args[1]), list(Path(tempfile.gettempdir()).iterdir()))

    monkeypatch.setitem(sys.modules, "webview", types.SimpleNamespace(
        settings=seen.settings, create_window=create_window, start=start))
    monkeypatch.setattr(app, "on_stop_signals", lambda close: setattr(seen, "on_stop", close))
    return seen


def test_window_gets_a_live_server_and_quit_cleans_up(tmp_dirs, monkeypatch):
    seen = {}

    def window(url, stop):
        seen.update(url=url, page=get(url), dirs=list(tmp_dirs.iterdir()))
        return True

    port = free_port()
    monkeypatch.setattr(app.web, "PORT", port)
    monkeypatch.setattr(app, "window", window)
    assert app.app() == 0
    assert seen["url"] == f"http://127.0.0.1:{port}/"
    status, page = seen["page"]
    assert status == 200 and "const APP = true;" in page, "страница не знает, что её запустило приложение"
    assert len(seen["dirs"]) == 1 and list(tmp_dirs.iterdir()) == [], "временный каталог не удалён"
    with pytest.raises(OSError):
        get(seen["url"])  # сервер остановлен


def test_window_failure_still_cleans_up(tmp_dirs, monkeypatch):
    def window(url, stop):
        raise RuntimeError("no display")

    monkeypatch.setattr(app.web, "PORT", free_port())
    monkeypatch.setattr(app, "window", window)
    with pytest.raises(RuntimeError):
        app.app()
    assert list(tmp_dirs.iterdir()) == []


def test_window_wiring(tmp_dirs, monkeypatch, webview):
    """Окно открывает страницу сервера, «Скачать .docx» и .zip сервер сохраняет через окно «Сохранить»;
    настройки страницы (тема, язык) не стираются между запусками; kill закрывает окно, а «Завершить»
    останавливает сервер и удаляет временные файлы до exit()."""
    real, servers = app.web.Server, []

    def server(*args, **kwargs):
        servers.append(real(*args, **kwargs))
        return servers[-1]

    port = free_port()
    monkeypatch.setattr(app.web, "PORT", port)
    monkeypatch.setattr(app.web, "Server", server)
    assert app.app() == 0
    (win,) = webview.windows
    assert win.args == ("md2docx", f"http://127.0.0.1:{port}/")
    assert win.kwargs["text_select"] and win.kwargs["background_color"] == "#0a0a0a"
    assert servers[-1].save_as is app.save_as, "«Скачать .docx» и .zip не откроют окно «Сохранить»"
    assert webview.settings == {}, "скачивание силами окна: WebKit не заменяет существующий файл"
    assert webview.started == {"private_mode": False, "storage_path": None}
    assert webview.on_stop is win.destroy, "kill и закрытый Терминал не закроют окно"
    assert webview.after_closing == (False, []), "после «Завершить» сервер работает или остались временные файлы"


def test_windows_keeps_window_data_in_its_own_folder(tmp_dirs, monkeypatch, webview):
    monkeypatch.setattr(app, "sys", types.SimpleNamespace(platform="win32"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_dirs))
    monkeypatch.setattr(app.web, "PORT", free_port())
    assert app.app() == 0
    assert webview.started == {"private_mode": False,
                               "storage_path": os.path.join(str(tmp_dirs), "md2docx", "webview")}


@pytest.mark.parametrize("chosen, path", [(("C:\\Users\\me\\a.docx",), "C:\\Users\\me\\a.docx"),
                                          ("/Volumes/Docs/a.docx", "/Volumes/Docs/a.docx"), (None, None), ((), None)],
                         ids=["windows", "macos", "cancel", "empty"])
def test_save_dialog_answer(monkeypatch, chosen, path):
    """Окно «Сохранить» pywebview: в Windows — кортеж из одного пути, в macOS — строка, «Отмена» — None."""
    asked = []
    win = types.SimpleNamespace(create_file_dialog=lambda *a, **kw: asked.append((a, kw)) or chosen)
    monkeypatch.setitem(sys.modules, "webview", types.SimpleNamespace(
        windows=[win], FileDialog=types.SimpleNamespace(SAVE=30)))
    assert app.save_as("a.docx") == path
    assert asked == [((30, os.path.join(os.path.expanduser("~"), "Downloads")), {"save_filename": "a.docx"})]


@pytest.mark.parametrize("lang, rtl", [("en", False), ("ar", True)])
def test_without_webview2_the_app_explains_and_exits(tmp_dirs, monkeypatch, webview, lang, rtl):
    """Без WebView2 pywebview берёт движок Internet Explorer: вместо пустого окна — сообщение и код 1.
    Сообщение — поверх других окон, на арабском — справа налево."""
    shown = []
    monkeypatch.setitem(sys.modules, "ctypes", types.SimpleNamespace(windll=types.SimpleNamespace(
        user32=types.SimpleNamespace(MessageBoxW=lambda *a: shown.append(a)))))
    monkeypatch.setattr(app, "language", lambda: lang)
    webview.renderer = "mshtml"
    monkeypatch.setattr(app.web, "PORT", free_port())
    assert app.app() == 1
    ((_, text, title, flags),) = shown
    assert app.WEBVIEW2 in text and title == "md2docx"
    assert flags & 0x10 and flags & 0x10000, "не значок ошибки или не на переднем плане"
    assert bool(flags & 0x100000) == bool(flags & 0x80000) == rtl
    assert list(tmp_dirs.iterdir()) == []


@pytest.mark.parametrize("renderer", ["cocoa", "edgechromium", "gtk", "qt"])
def test_other_renderers_are_supported(renderer):
    assert app.supported(renderer) is True


@pytest.mark.skipif(sys.platform == "win32", reason="os.kill под Windows завершает процесс, а не шлёт сигнал")
@pytest.mark.parametrize("sig", app.STOP_SIGNALS, ids=lambda s: signal.Signals(s).name)
def test_kill_closes_the_window(sig):
    """kill (SIGTERM) и закрытый Терминал (SIGHUP) закрывают окно, хотя главный поток занят циклом окна:
    здесь он в коде на C, и обработчик сигнала на Python сработал бы только после его возврата."""
    closed, sent = [], []
    app.on_stop_signals(lambda: closed.append(time.monotonic()))
    timer = threading.Timer(0.2, lambda: sent.append(time.monotonic()) or os.kill(os.getpid(), sig))
    timer.start()
    ctypes.CDLL(None).system(b"sleep 2")  # цикл окна — тоже код на C
    timer.join()
    assert closed and closed[0] - sent[0] < 1, "сигнал не закрыл окно, пока главный поток в цикле окна"


@pytest.mark.skipif(sys.platform == "win32", reason="os.kill под Windows завершает процесс, а не шлёт сигнал")
def test_every_kill_closes_the_window():
    """Окно не показалось (pywebview ждёт его 20 с и бросает исключение) — следующий kill пробует снова."""
    closes = []

    def close():
        closes.append(len(closes))
        if len(closes) == 1:
            raise RuntimeError("Main window failed to start")

    app.on_stop_signals(close)
    for n in (1, 2):
        os.kill(os.getpid(), signal.SIGTERM)
        deadline = time.monotonic() + 10
        while len(closes) < n and time.monotonic() < deadline:
            time.sleep(0.01)
    assert closes == [0, 1], "второй kill не закрыл окно"


def test_second_launch_gets_its_own_window(tmp_dirs, monkeypatch):
    """Порт занят первым экземпляром — второй поднимает свой сервер на любом свободном порту и своё окно."""
    first = app.web.Server(free_port(), app=True)
    threading.Thread(target=first.serve_forever, daemon=True).start()
    urls = []
    try:
        monkeypatch.setattr(app.web, "PORT", first.port)
        monkeypatch.setattr(app, "window", lambda url, stop: urls.append(url) or True)
        assert app.app() == 0
        (url,) = urls
        assert url.startswith("http://127.0.0.1:") and not url.endswith(f":{first.port}/")
        assert list(tmp_dirs.iterdir()) == [first.root], "второй запуск не убрал за собой"
    finally:
        first.shutdown()
        first.server_close()
        first.cleanup()


def test_port_held_by_another_program_gives_any_free_port(tmp_dirs, monkeypatch):
    urls = []
    with socket.socket() as other:
        other.bind(("127.0.0.1", 0))
        other.listen()
        port = other.getsockname()[1]
        monkeypatch.setattr(app.web, "PORT", port)
        monkeypatch.setattr(app, "window", lambda url, stop: urls.append(url) or True)
        assert app.app() == 0
    (url,) = urls
    assert url.startswith("http://127.0.0.1:") and not url.endswith(f":{port}/")


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
