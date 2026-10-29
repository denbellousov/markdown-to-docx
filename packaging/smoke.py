#!/usr/bin/env python3
"""Дымовой тест запущенного md2docx: тот же путь, что проходит страница в браузере.

    python packaging/smoke.py [base_url] [sample.md]

base_url — адрес md2docx (по умолчанию http://127.0.0.1:51536/), sample.md — что конвертировать
(по умолчанию samples/08-showcase.md). До 60 с ждёт ответа сервера, затем: токен со страницы →
задание → загрузка файла → конвертация → скачивание .docx и проверка, что это zip
с word/document.xml. Заодно проверяет, что в сборку попали переводы интерфейса (locales/).

Код выхода 0 — всё прошло, иначе 1 и строка «smoke: FAIL: …» с причиной.
Только стандартная библиотека: работает на любом Python 3.10+ без установки пакетов.
"""

from __future__ import annotations

import http.client
import io
import json
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from pathlib import Path

BASE = "http://127.0.0.1:51536/"
SAMPLE = Path(__file__).resolve().parent.parent / "samples" / "08-showcase.md"
WAIT = 60      # с: сколько ждать, пока приложение поднимет сервер
TIMEOUT = 120  # с: на один запрос — конвертация на холодном раннере бывает небыстрой

# напрямую: прокси из переменных окружения раннера к 127.0.0.1 отношения не имеет
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def fail(message):
    sys.exit(f"smoke: FAIL: {message}")


def reason(exc):
    return getattr(exc, "reason", None) or exc


def request(method, url, body=None, headers=None):
    """Тело ответа 2xx; иначе — провал с текстом ошибки сервера (он отвечает {"error": …})."""
    where = f"{method} {urllib.parse.urlsplit(url).path}"
    req = urllib.request.Request(url, data=body, headers=headers or {}, method=method)
    try:
        with OPENER.open(req, timeout=TIMEOUT) as resp:
            return resp.read()
    except urllib.error.HTTPError as exc:
        text = exc.read().decode("utf-8", "replace")
        try:
            text = json.loads(text)["error"]
        except (ValueError, KeyError, TypeError):
            pass
        fail(f"{where} -> HTTP {exc.code}: {str(text).strip()[:500]}")
    except (OSError, http.client.HTTPException) as exc:  # сервер упал или оборвал соединение
        fail(f"{where}: {reason(exc)}")


def request_json(method, url, body=None, headers=None):
    raw = request(method, url, body, headers)
    try:
        answer = json.loads(raw)
    except ValueError:
        answer = None
    if not isinstance(answer, dict):
        fail(f"{method} {urllib.parse.urlsplit(url).path}: expected a JSON object, got {raw[:200]!r}")
    return answer


def wait_for_page(base):
    """HTML страницы интерфейса; приложение поднимает сервер не сразу."""
    deadline = time.monotonic() + WAIT
    while True:
        try:
            with OPENER.open(base, timeout=10) as resp:
                return resp.read().decode("utf-8")
        except urllib.error.HTTPError as exc:  # отвечает, но не так — ждать дальше незачем
            fail(f"GET {base} -> HTTP {exc.code}")
        except (OSError, http.client.HTTPException) as exc:  # порт ещё не слушают
            if time.monotonic() > deadline:
                fail(f"no answer from {base} within {WAIT} s: {reason(exc)}")
        time.sleep(1)


def main(argv):
    for stream in (sys.stdout, sys.stderr):  # консоль Windows и лог CI: без UnicodeEncodeError
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="backslashreplace")
    base = argv[1] if len(argv) > 1 else BASE
    base = base if base.endswith("/") else base + "/"
    sample = Path(argv[2]) if len(argv) > 2 else SAMPLE
    try:
        data = sample.read_bytes()
    except OSError as exc:
        fail(f"cannot read the sample {sample}: {exc.strerror or exc}")

    page = wait_for_page(base)
    token = re.search(r'const TOKEN = "([^"]+)";', page)
    if not token:
        fail(f'no `const TOKEN = "...";` on the page at {base} - is this md2docx?')
    auth = {"X-Token": token[1]}

    # переводы страницы: если папка locales не попала в сборку, каталоги всех языков пусты
    i18n = re.search(r"const I18N = ", page)
    try:
        catalogs = json.JSONDecoder().raw_decode(page, i18n.end())[0] if i18n else None
    except ValueError:
        catalogs = None
    if not isinstance(catalogs, dict) or not any(catalogs.values()):
        fail("the page has no interface translations (const I18N): locales/ is missing from the build")

    job = request_json("POST", base + "api/jobs", b"", auth).get("job")
    if not job:
        fail("POST /api/jobs: no job id in the answer")
    request("PUT", f"{base}api/jobs/{job}/files?path={urllib.parse.quote(sample.name)}", data,
            {**auth, "Content-Type": "application/octet-stream", "Content-Length": str(len(data))})
    answer = request_json("POST", f"{base}api/jobs/{job}/convert",
                          json.dumps({"orientation": "auto"}).encode(),
                          {**auth, "Content-Type": "application/json"})
    first = (answer.get("results") or [{}])[0]
    if first.get("ok") is not True or not first.get("url"):
        fail(f"conversion of {sample.name} failed: {first.get('error') or answer}")

    docx = request("GET", urllib.parse.urljoin(base, first["url"]), None, auth)
    try:
        with zipfile.ZipFile(io.BytesIO(docx)) as z:
            names, broken = z.namelist(), z.testzip()
    except zipfile.BadZipFile as exc:
        fail(f"the download is not a .docx (zip): {exc}; it starts with {docx[:8]!r}")
    if "word/document.xml" not in names:
        fail(f"the .docx has no word/document.xml; parts: {', '.join(names[:10])}")
    if broken:
        fail(f"the .docx is damaged: bad CRC in {broken}")
    warnings = first.get("warnings") or []
    print(f"smoke: OK: {first.get('name')}, {len(docx):,} bytes, {len(names)} parts, "
          f"{len(warnings)} warning(s); {first.get('summary') or 'no summary'}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
