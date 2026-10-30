#!/usr/bin/env python3
"""md2docx_app — точка входа приложений md2docx для macOS и Windows.

Без ключей: запускает интерфейс md2docx_web.py и показывает его в собственном окне — pywebview
(WebKit на macOS, WebView2 на Windows), без браузера. Выход (закрытие окна, меню, Dock, kill)
останавливает сервер и удаляет временные файлы. С ключами работает как md2docx_web.py
(--port N, --no-browser): так приложение запускают из терминала и дымовой тест сборки.

Приложения собирает PyInstaller по packaging/md2docx.spec.
"""

from __future__ import annotations

import os
import signal
import socket
import sys
import threading

if sys.stdout is None or sys.stderr is None:  # оконная сборка под Windows: потоков вывода нет
    sys.stdout = sys.stdout or open(os.devnull, "w", encoding="utf-8")
    sys.stderr = sys.stderr or open(os.devnull, "w", encoding="utf-8")

import md2docx_web as web  # noqa: E402
from md2docx import RTL, language, tr  # noqa: E402

STOP_SIGNALS = [getattr(signal, n) for n in ("SIGTERM", "SIGHUP") if hasattr(signal, n)]  # kill, закрыли Терминал
WEBVIEW2 = "https://developer.microsoft.com/microsoft-edge/webview2/consumer/"


def on_stop_signals(close):
    """kill и закрытый Терминал вызывают close() — как закрытие окна.

    Обработчик Python выполнился бы, только когда главный поток вернётся в Python, а он стоит
    в цикле окна; номер сигнала в wakeup_fd пишется сразу — его ждёт отдельный поток."""
    wake, wakeup = socket.socketpair()
    wakeup.setblocking(False)
    signal.set_wakeup_fd(wakeup.fileno())
    for s in STOP_SIGNALS:
        signal.signal(s, lambda *_: None)

    def wait(keep):  # keep — wakeup: пока жив поток, дескриптор из set_wakeup_fd не закрыт
        while wake.recv(1):  # поток живёт до выхода: каждый сигнал снова закрывает окно
            try:
                close()
            except Exception:  # окно так и не показалось (pywebview ждёт его 20 с) — ждём следующий сигнал
                pass

    threading.Thread(target=wait, args=(wakeup,), daemon=True).start()


def supported(renderer):
    """Страница работает в WebView2; без него pywebview берёт движок Internet Explorer, и окно пустое."""
    if renderer != "mshtml":
        return True
    import ctypes  # mshtml бывает только под Windows
    flags = 0x10 | 0x10000  # MB_ICONERROR | MB_SETFOREGROUND: окна нет — сообщение не должно уйти под чужие
    if language() in RTL:
        flags |= 0x100000 | 0x80000  # MB_RTLREADING | MB_RIGHT
    ctypes.windll.user32.MessageBoxW(
        None, tr("md2docx needs Microsoft Edge WebView2 Runtime to show its window. "
                 "Install it from {url} and open md2docx again.", url=WEBVIEW2), "md2docx", flags)
    return False


def save_as(name):
    """Окно «Сохранить» для «Скачать .docx» и .zip: выбранный путь или None («Отмена»)."""
    import webview

    chosen = webview.windows[0].create_file_dialog(
        webview.FileDialog.SAVE, os.path.join(os.path.expanduser("~"), "Downloads"), save_filename=name)
    if isinstance(chosen, (tuple, list)):  # Windows отдаёт кортеж из одного пути, macOS — строку
        chosen = chosen[0] if chosen else None
    return chosen or None


def window(url, stop):
    """Окно со страницей url; возвращается, когда его закрыли. False — окно показать нечем."""
    import webview

    win = webview.create_window("md2docx", url, width=900, height=780, min_size=(420, 480),
                                background_color="#0a0a0a", text_select=True, zoomable=True)
    shown = True

    def initialized(renderer):
        nonlocal shown
        shown = supported(renderer)
        return shown

    win.events.initialized += initialized
    win.events.closing += stop  # «Завершить» в меню и Dock выходит через exit(): finally в app() не наступит
    on_stop_signals(win.destroy)
    # тема и язык, выбранные на странице, сохраняются между запусками; на Windows — в своей папке,
    # а не в общей для всех приложений на pywebview (на macOS данные и так у каждого приложения свои)
    local = os.environ.get("LOCALAPPDATA") if sys.platform == "win32" else None
    webview.start(private_mode=False, storage_path=local and os.path.join(local, "md2docx", "webview"))
    return shown


def app():
    try:
        server = web.Server(web.PORT, app=True, save_as=save_as)
    except OSError:  # порт занят — второй запуск или чужая программа: окну подходит любой свободный
        server = web.Server(0, app=True, save_as=save_as)
    # poll_interval — сколько shutdown() ждёт цикл сервера, а его зовёт и закрытие окна
    threading.Thread(target=server.serve_forever, kwargs={"poll_interval": 0.1}, daemon=True).start()

    def stop():
        server.shutdown()
        server.server_close()
        server.cleanup()

    try:
        shown = window(f"http://{web.HOST}:{server.port}/", stop)
    finally:
        stop()
    return 0 if shown else 1


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    argv = [a for a in argv if not a.startswith("-psn_")]  # старые macOS добавляют при запуске из Finder
    # без ключей — окно: файл, брошенный на значок в Windows, приходит аргументом, а ошибку argparse не видно
    return web.main(argv, prog="md2docx") if any(a.startswith("-") for a in argv) else app()


if __name__ == "__main__":
    sys.exit(main())
