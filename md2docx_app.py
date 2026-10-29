#!/usr/bin/env python3
"""md2docx_app — точка входа приложений md2docx для macOS и Windows.

Без ключей: запускает интерфейс md2docx_web.py, открывает его в браузере и показывает
маленькое окно с адресом и кнопками «Открыть в браузере» и «Выйти». Выход (закрытие окна,
kill) останавливает сервер и удаляет временные файлы. С ключами работает как md2docx_web.py
(--port N, --no-browser): так приложение запускают из терминала и дымовой тест сборки.

Приложения собирает PyInstaller по packaging/md2docx.spec.
"""

from __future__ import annotations

import http.client
import os
import signal
import sys
import threading
import webbrowser

if sys.stdout is None or sys.stderr is None:  # оконная сборка под Windows: потоков вывода нет
    sys.stdout = sys.stdout or open(os.devnull, "w", encoding="utf-8")
    sys.stderr = sys.stderr or open(os.devnull, "w", encoding="utf-8")

import md2docx_web as web  # noqa: E402
from md2docx import tr  # noqa: E402

ICON = web.HERE / "packaging" / "icon.png"
STOP_SIGNALS = [getattr(signal, n) for n in ("SIGTERM", "SIGHUP") if hasattr(signal, n)]  # kill, закрыли Терминал


def running_here(port):
    """На порту уже отвечает md2docx: второй запуск приложения или md2docx_web.py."""
    conn = http.client.HTTPConnection(web.HOST, port, timeout=2)
    try:
        conn.request("GET", "/")
        return b"md2docx v" in conn.getresponse().read()
    except OSError:
        return False
    finally:
        conn.close()


def window(url):
    """Окно с адресом; возвращается, когда его закрыли или выбрали «Выйти»."""
    import tkinter as tk
    from tkinter import font, ttk

    rtl = web.md2docx.language() == "ar"
    side, anchor = ("right", "e") if rtl else ("left", "w")
    saved = {s: signal.getsignal(s) for s in STOP_SIGNALS}
    root = tk.Tk()
    for s, handler in saved.items():  # Tk() ставит свои: выход с кодом 1 без очистки
        signal.signal(s, handler)
    root.title("md2docx")
    root.resizable(False, False)
    if ICON.exists():  # иначе в панели задач Windows — перо Tk
        root.iconphoto(True, tk.PhotoImage(file=str(ICON)))
    bold = font.nametofont("TkDefaultFont").copy()
    bold.configure(weight="bold")

    def open_browser(*_):
        webbrowser.open(url)

    def reopen():  # macOS: щелчок по значку в Dock, когда приложение уже запущено
        root.deiconify()
        root.lift()
        open_browser()

    box = ttk.Frame(root, padding=20)
    box.pack(fill="both", expand=True)
    ttk.Label(box, text=f"md2docx v{web.md2docx.__version__}", font=bold).pack(anchor=anchor)
    ttk.Label(box, text=tr("The interface is open in the browser at this address:")).pack(anchor=anchor, pady=(8, 4))
    address = ttk.Entry(box, width=len(url) + 2)
    address.insert(0, url)
    address.state(["readonly"])
    address.pack(anchor=anchor, fill="x")
    buttons = ttk.Frame(box)
    buttons.pack(anchor=anchor, pady=(14, 0))
    ttk.Button(buttons, text=tr("Open in browser"), command=open_browser, default="active").pack(side=side)
    ttk.Button(buttons, text=tr("Quit"), command=root.destroy).pack(side=side, padx=8)
    ttk.Label(box, text=tr("Closing this window stops md2docx and deletes its temporary files."),
              wraplength=380, justify=side).pack(anchor=anchor, pady=(14, 0))

    root.bind("<Return>", open_browser)
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    if sys.platform == "darwin":
        root.createcommand("::tk::mac::ReopenApplication", reopen)
        root.createcommand("::tk::mac::Quit", root.destroy)
    root.update_idletasks()
    x = (root.winfo_screenwidth() - root.winfo_reqwidth()) // 2
    y = (root.winfo_screenheight() - root.winfo_reqheight()) // 3
    root.geometry(f"+{x}+{y}")
    root.after(0, open_browser)
    root.mainloop()


def app():
    try:
        server = web.Server(web.PORT, app=True)
    except OSError:
        if running_here(web.PORT):  # уже запущен — показать его, второй сервер не нужен
            webbrowser.open(f"http://{web.HOST}:{web.PORT}/")
            return 0
        server = web.Server(0, app=True)  # порт занят чужой программой: любой свободный, адрес — в окне
    threading.Thread(target=server.serve_forever, daemon=True).start()
    for s in STOP_SIGNALS:  # как «Выйти»
        signal.signal(s, web._stop)
    try:
        window(f"http://{web.HOST}:{server.port}/")
    finally:
        server.shutdown()
        server.server_close()
        server.cleanup()
    return 0


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    argv = [a for a in argv if not a.startswith("-psn_")]  # старые macOS добавляют при запуске из Finder
    # без ключей — окно: файл, брошенный на значок в Windows, приходит аргументом, а ошибку argparse не видно
    return web.main(argv, prog="md2docx") if any(a.startswith("-") for a in argv) else app()


if __name__ == "__main__":
    sys.exit(main())
