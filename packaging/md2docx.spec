# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 6.22.3: приложение md2docx, точка входа — md2docx_app.py, окно — pywebview 6.2.1
# (его файлы — js, на Windows ещё lib — собирает собственный хук pywebview).
#
# Сборка из корня репозитория:  pyinstaller --noconfirm packaging/md2docx.spec
#   macOS   → dist/md2docx.app — оконное, onedir внутри пакета .app;
#   Windows → dist/md2docx.exe — оконное, один файл.
# Значок packaging/icon.png PyInstaller сам переводит в .icns и .ico через Pillow
# (Pillow и так ставится из requirements.txt).
import os
import sys
import sysconfig

ROOT = os.path.dirname(SPECPATH)  # SPECPATH — абсолютный путь к каталогу этого файла
sys.path.insert(0, ROOT)
import md2docx  # noqa: E402 — версия приложения: md2docx.__version__

ICON = os.path.join(SPECPATH, "icon.png")

a = Analysis(
    [os.path.join(ROOT, "md2docx_app.py")],
    datas=[
        # переводы md2docx ищет рядом со своим модулем, то есть в sys._MEIPASS/locales
        (os.path.join(ROOT, "locales"), "locales"),
    ],
)
pyz = PYZ(a.pure)

if sys.platform == "darwin":
    # Минимальная macOS — 11.3 или выше, если её требует Python, которым собираем (python.org 3.13:
    # 10.13): окно pywebview на каждой навигации спрашивает WKNavigationAction.shouldPerformDownload,
    # а он появился в macOS 11.3.
    min_macos = max(sysconfig.get_config_var("MACOSX_DEPLOYMENT_TARGET") or "10.13", "11.3",
                    key=lambda v: tuple(int(n) for n in v.split(".")))
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name="md2docx",
        console=False,
        argv_emulation=False,  # файлы, брошенные на значок, приложению не нужны; -psn_ оно отбрасывает само
        upx=False,
    )
    coll = COLLECT(exe, a.binaries, a.datas, name="md2docx", upx=False)
    app = BUNDLE(
        coll,
        name="md2docx.app",
        icon=ICON,
        bundle_identifier="io.github.denbellousov.md2docx",
        version=md2docx.__version__,
        info_plist={"NSHighResolutionCapable": True, "LSMinimumSystemVersion": min_macos},
    )
else:
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        [],
        name="md2docx",
        console=False,
        icon=ICON,
        upx=False,
    )
