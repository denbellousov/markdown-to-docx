"""Тесты локального интерфейса md2docx_web.py: сервер в потоке на свободном порту, запросы —
те же, что шлёт страница. Результат сверяется с CLI: word/document.xml байт в байт,
предупреждения — дословно. Остановка, занятый порт и запуск .command — отдельным процессом.

    python3 -m pytest tests/test_web.py -q --basetemp=<свой каталог> -p no:cacheprovider
"""

from __future__ import annotations

import http.client
import io
import json
import os
import re
import shutil
import signal
import socket
import struct
import subprocess
import sys
import threading
import zipfile
import zlib
from html.parser import HTMLParser
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import quote

import pytest

ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "samples"
WEB = ROOT / "md2docx_web.py"
sys.path.insert(0, str(ROOT))
import md2docx_web as web  # noqa: E402

SAMPLE_FILES = sorted(SAMPLES.glob("*.md"))
RESAVED = ROOT / "tests" / "fixtures" / "02-tables-resaved.docx"  # пересохранён в Word для Mac
CLI_FLAGS = {"auto": [], "landscape": ["--landscape"], "no-auto": ["--no-auto-landscape"]}
WARN = "Внимание: "
LOG_LINE = re.compile(r"\d\d:\d\d:\d\d GMT[+-]\d{1,2}(:\d\d)?  (GET|POST|PUT) /[^\s?]* \d{3}")

# --------------------------------------------------------------------------
# Сервер и запросы
# --------------------------------------------------------------------------


@pytest.fixture
def srv(tmp_path):
    server = web.Server(0, tmp_parent=tmp_path)  # 0 — свободный порт от ОС
    threading.Thread(target=server.serve_forever, kwargs={"poll_interval": 0.05}, daemon=True).start()
    yield server
    server.shutdown()
    server.server_close()
    server.cleanup()
    assert not server.root.exists(), "временный каталог не удалён"


def call(srv, method, path, body=None, *, token=True, headers=None):
    """Запрос → (статус, заголовки, тело). Host — 127.0.0.1:<порт>, как у браузера."""
    conn = http.client.HTTPConnection("127.0.0.1", srv.port, timeout=120)
    hdrs = {"X-Token": srv.token} if token else {}
    hdrs.update(headers or {})
    try:
        conn.request(method, path, body=body, headers=hdrs)
        resp = conn.getresponse()
        return resp.status, resp.headers, resp.read()
    finally:
        conn.close()


def new_job(srv):
    status, _, body = call(srv, "POST", "/api/jobs")
    assert status == 201, body
    return json.loads(body)["job"]


def upload(srv, job, rel, data, **kw):
    return call(srv, "PUT", f"/api/jobs/{job}/files?path={quote(rel, safe='')}", data, **kw)


def run(srv, job, orientation="auto"):
    status, _, body = call(srv, "POST", f"/api/jobs/{job}/convert",
                           json.dumps({"orientation": orientation}).encode(),
                           headers={"Content-Type": "application/json"})
    assert status == 200, body.decode()
    return json.loads(body)


def fetch(srv, url):
    status, headers, body = call(srv, "GET", url)
    assert status == 200, body[:500]
    return headers, body


def stored(srv):
    return sorted(p.relative_to(srv.root) for p in srv.root.rglob("*") if p.is_file())


def document_xml(data: bytes) -> bytes:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return z.read("word/document.xml")


def members(data: bytes) -> dict:
    """Части docx: имя → байты (без сжатия)."""
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return {name: z.read(name) for name in z.namelist()}


def cli(*args):
    return subprocess.run([sys.executable, str(ROOT / "md2docx.py"), *map(str, args)],
                          capture_output=True, text=True, timeout=300)


def cli_warnings(proc):
    return [line[len(WARN):] for line in proc.stderr.splitlines() if line.startswith(WARN)]


def client_comments(page):
    """Комментарии в разметке, стилях и скриптах страницы; строковые литералы JS не в счёт."""
    found = re.findall(r"<!--.*?-->", page, re.S)
    for css in re.findall(r"<style>(.*?)</style>", page, re.S):
        found += re.findall(r"/\*.*?\*/", css, re.S)
    for js in re.findall(r"<script[^>]*>(.*?)</script>", page, re.S):
        code = re.sub(r'"(?:[^"\\\n]|\\.)*"', '""', js)
        found += re.findall(r"/\*.*?\*/|//[^\n]*", code, re.S)
    return found


def css_rules(page):
    """Правила из <style> страницы → [(селекторы, объявления без пробелов)]."""
    css = "".join(re.findall(r"<style>(.*?)</style>", page, re.S))
    return [([s.strip() for s in sel.split(",")], re.sub(r"\s+", "", body))
            for sel, body in re.findall(r"([^{}]+)\{([^{}]*)\}", css)]


class HiddenById(HTMLParser):
    """Скрыт ли в разметке элемент с данным id атрибутом hidden — своим или предка."""
    VOID = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source",
            "track", "wbr"}

    def __init__(self, target):
        super().__init__()
        self.target, self.stack, self.hidden = target, [], None

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        hidden = "hidden" in attrs or bool(self.stack and self.stack[-1][1])
        if attrs.get("id") == self.target:
            self.hidden = hidden
        if tag not in self.VOID:
            self.stack.append((tag, hidden))

    def handle_endtag(self, tag):
        while self.stack and self.stack.pop()[0] != tag:
            pass


def initially_hidden(page, element_id):
    parser = HiddenById(element_id)
    parser.feed(page)
    assert parser.hidden is not None, f"нет элемента #{element_id}"
    return parser.hidden


def tiny_png() -> bytes:
    """PNG 1×1 без Pillow."""
    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(b"\x00\xff\x00\x00")) + chunk(b"IEND", b""))


# --------------------------------------------------------------------------
# Страница
# --------------------------------------------------------------------------


def test_page_dark_by_default_with_token_and_version(srv):
    status, headers, body = call(srv, "GET", "/", token=False)
    page = body.decode()
    assert status == 200 and headers["Content-Type"] == "text/html; charset=utf-8"
    assert f'const TOKEN = "{srv.token}"' in page
    assert f"md2docx v{web.md2docx.__version__}" in page
    assert '<html lang="ru" data-theme="dark">' in page
    for label in ("Перетащите .md или .docx сюда", "исправить шрифты для Windows",
                  "Из папки .docx не берутся", "Выбрать файлы", "Авто", "Всё альбомное",
                  "Без автоальбомной", "готово", "ошибка", "шрифты исправлены",
                  "исправлять нечего — подсказки на месте", "Скачать .docx", "Скачать всё (.zip)"):
        assert label in page, label
    assert 'accept=".md,.markdown,.docx,image/*"' in page
    assert client_comments(page) == [], "в клиентском коде — без комментариев"
    for planted in ("<!-- x -->", "<style>a{}/* x */</style>", "<script>f(); // x\n</script>",
                    '<script>g("//"); /* x */</script>'):
        assert client_comments(planted), planted  # проверка умеет краснеть
    assert not re.search(r"__[A-Z_]+__", page.replace(srv.token, "")), "подстановка не сработала"
    assert re.search(r"script-src 'nonce-[\w-]+'", headers["Content-Security-Policy"])


def test_zip_button_hidden_until_server_offers_zip(srv):
    """«Скачать всё (.zip)» в разметке скрыта атрибутом hidden (у её блока), и он действительно
    скрывает: правило [hidden]{display:none!important}. Без него .actions{display:flex} перебивало
    hidden, и кнопка была видна всегда — даже до выбора файлов. Показывает кнопку страница
    только по полю zip ответа сервера; правило «zip — когда скачать можно хотя бы два файла»
    проверяют тесты сервера (test_zip_of_all_results и др.). Скрипт страницы тесты не исполняют:
    здесь — разметка, CSS и то, что видимость взята из ответа сервера, а не посчитана заново."""
    def hides(html_text):
        return any("[hidden]" in selectors and re.search(r"(^|;)display:none!important(;|$)", body)
                   for selectors, body in css_rules(html_text))

    page = call(srv, "GET", "/", token=False)[2].decode()
    assert initially_hidden(page, "zip"), "кнопка zip в разметке не скрыта"
    assert hides(page), "нет правила [hidden]{display:none!important}"
    assert "zipUrl = data.zip;" in page and "actions.hidden = !zipUrl;" in page, \
        "видимость кнопки zip — не по полю zip ответа сервера"
    # проверка умеет краснеть: без !important правило не в счёт, hidden соседа — не у кнопки
    assert not hides("<style>.actions{display:flex}[hidden]{display:none}</style>")
    assert hides("<style>.actions{display:flex}\n[hidden] { display: none !important }</style>")
    assert not initially_hidden('<div class="actions"><a id="zip">z</a></div><p hidden></p>', "zip")


# --------------------------------------------------------------------------
# Конвертация: то же, что CLI
# --------------------------------------------------------------------------


def test_samples_match_cli_in_every_orientation(srv, tmp_path):
    """Одна загрузка образцов, три ориентации подряд без повторной загрузки: каждый docx совпадает
    с CLI по word/document.xml, предупреждения и строка итога — как в выводе CLI."""
    assert SAMPLE_FILES, "нет samples/*.md"
    job = new_job(srv)
    for md in SAMPLE_FILES:
        status, _, body = upload(srv, job, md.name, md.read_bytes())
        assert status == 201, body
    xml = {}
    for mode, flags in CLI_FLAGS.items():
        cli_dir = tmp_path / f"cli-{mode}"
        cli_dir.mkdir()
        proc = cli(*[shutil.copy(md, cli_dir) for md in SAMPLE_FILES], *flags)
        assert proc.returncode == 0, proc.stderr
        data = run(srv, job, mode)
        assert data["orientation"] == mode
        assert [r["source"] for r in data["results"]] == [md.name for md in SAMPLE_FILES]
        assert [w for r in data["results"] for w in r["warnings"]] == cli_warnings(proc)
        assert [r["summary"] for r in data["results"]] == [
            line.rsplit(" — ", 1)[1] for line in proc.stdout.splitlines() if line.startswith("Готово: ")]
        for r in data["results"]:
            assert r["ok"], r
            headers, docx = fetch(srv, r["url"])
            assert headers["Content-Type"] == web.DOCX
            xml[mode, r["source"]] = document_xml(docx)
            assert xml[mode, r["source"]] == document_xml((cli_dir / r["name"]).read_bytes()), \
                f"{mode}: {r['source']}"
    # проба различает: переключатель действительно меняет документ
    assert xml["auto", "04-auto-landscape.md"] != xml["no-auto", "04-auto-landscape.md"]
    assert xml["auto", "01-headings.md"] != xml["landscape", "01-headings.md"]
    assert len(list((srv.root / job).glob("run*"))) == 1, "прежние прогоны не удалены"


def test_converter_warning_reaches_response_verbatim(srv, tmp_path):
    md = tmp_path / "ссылки.md"
    md.write_text("# Ссылки\n\nСм. [локальный файл](file:///etc/hosts) и [сайт](https://example.com).\n",
                  encoding="utf-8")
    proc = cli(md, "-o", tmp_path / "ссылки.docx")
    assert proc.returncode == 0, proc.stderr
    expected = cli_warnings(proc)
    assert len(expected) == 1 and "file:///etc/hosts" in expected[0]
    job = new_job(srv)
    assert upload(srv, job, md.name, md.read_bytes())[0] == 201
    (r,) = run(srv, job)["results"]
    assert r["ok"] and r["warnings"] == expected


def test_folder_keeps_relative_images(srv):
    md = "# Отчёт\n\n![Схема](картинки/схема.png)\n".encode()
    job = new_job(srv)
    assert upload(srv, job, "отчёт/картинки/схема.png", tiny_png())[0] == 201
    assert upload(srv, job, "отчёт/док.md", md)[0] == 201
    (r,) = run(srv, job)["results"]
    assert r["ok"] and r["source"] == "отчёт/док.md" and r["name"] == "док.docx"
    assert r["warnings"] == []
    headers, docx = fetch(srv, r["url"])
    assert f"filename*=UTF-8''{quote('док.docx')}" in headers["Content-Disposition"]
    with zipfile.ZipFile(io.BytesIO(docx)) as z:
        assert any(name.startswith("word/media/") for name in z.namelist())
    # тот же .md без папки: картинки нет, и конвертер об этом предупреждает
    job = new_job(srv)
    assert upload(srv, job, "док.md", md)[0] == 201
    (r,) = run(srv, job)["results"]
    assert r["ok"] and any("картинка не найдена" in w for w in r["warnings"]), r["warnings"]


def test_zip_of_all_results(srv):
    job = new_job(srv)
    picked = SAMPLE_FILES[:3]
    for md in picked:
        assert upload(srv, job, md.name, md.read_bytes())[0] == 201
    assert upload(srv, job, "кодировка.md", "Текст в cp1251".encode("cp1251"))[0] == 201
    data = run(srv, job)
    ok = [r for r in data["results"] if r["ok"]]
    bad = [r for r in data["results"] if not r["ok"]]
    assert [r["source"] for r in ok] == [md.name for md in picked]
    assert [(r["source"], r["error"]) for r in bad] == [("кодировка.md", "файл не в кодировке UTF-8")]
    assert "url" not in bad[0]
    headers, archive = fetch(srv, data["zip"])
    assert headers["Content-Type"] == "application/zip"
    with zipfile.ZipFile(io.BytesIO(archive)) as z:
        assert z.testzip() is None
        assert z.namelist() == [r["name"] for r in ok]
        for r in ok:
            assert z.read(r["name"]) == fetch(srv, r["url"])[1]
    job = new_job(srv)  # один файл — архив не нужен
    assert upload(srv, job, picked[0].name, picked[0].read_bytes())[0] == 201
    assert run(srv, job)["zip"] is None


def test_duplicates_bad_orientation_and_no_md(srv):
    job = new_job(srv)
    assert upload(srv, job, "a.md", b"# 1\n")[0] == 201
    assert upload(srv, job, "a.md", b"# 2\n")[0] == 409
    assert upload(srv, job, "a.md/b.md", b"# 3\n")[0] == 409  # «a.md» — уже файл, не каталог
    assert (srv.root / job / "src" / "a.md").read_bytes() == b"# 1\n", "первый файл затёрт"
    assert call(srv, "POST", f"/api/jobs/{job}/convert", b'{"orientation": "sideways"}')[0] == 400
    assert call(srv, "POST", "/api/jobs/0123456789abcdef/convert", b"{}")[0] == 404
    job = new_job(srv)
    assert upload(srv, job, "схема.png", tiny_png())[0] == 201
    assert call(srv, "POST", f"/api/jobs/{job}/convert", b"{}")[0] == 400


def test_new_job_removes_previous_ones(srv):
    """Страница показывает только последнее задание: новое удаляет прежние с диска целиком,
    запрос к удалённому — 404 с понятным текстом."""
    first = new_job(srv)
    for name in ("a.md", "b.md"):
        assert upload(srv, first, name, b"# A\n")[0] == 201
    data = run(srv, first)
    assert data["zip"] and [p.name for p in srv.root.iterdir()] == [first]
    for n in range(3):
        job = new_job(srv)
        assert upload(srv, job, f"{n}.md", b"# N\n")[0] == 201
        run(srv, job)
        assert [p.name for p in srv.root.iterdir()] == [job], "на диске не одно задание"
    for method, url, body in (("GET", data["results"][0]["url"], None), ("GET", data["zip"], None),
                              ("POST", f"/api/jobs/{first}/convert", b"{}"),
                              ("PUT", f"/api/jobs/{first}/files?path=c.md", b"# C\n")):
        status, _, answer = call(srv, method, url, body)
        assert (status, json.loads(answer)) == (404, {"error": web.GONE}), url
    assert [p.name for p in srv.root.iterdir()] == [job], "запрос к удалённому заданию что-то создал"


# --------------------------------------------------------------------------
# .docx: исправление шрифтов, как md2docx.py --fix-fonts
# --------------------------------------------------------------------------


def test_docx_fonts_fixed_like_cli(srv, tmp_path, monkeypatch):
    """Пересохранённый в Word для Mac .docx: подсказки шрифтов возвращены в копию так же, как
    md2docx.py --fix-fonts на такой же копии; исходник фикстуры и загруженный файл не меняются,
    другая ориентация .docx заново не обрабатывает."""
    original = RESAVED.read_bytes()
    copy = tmp_path / RESAVED.name
    copy.write_bytes(original)
    proc = cli("--fix-fonts", copy)
    assert proc.returncode == 0, proc.stderr
    calls, real = [], web.md2docx.fix_fonts
    monkeypatch.setattr(web.md2docx, "fix_fonts", lambda path, **kw: calls.append(path) or real(path, **kw))
    job = new_job(srv)
    assert upload(srv, job, RESAVED.name, original)[0] == 201
    data = run(srv, job)
    (r,) = data["results"]
    assert (r["ok"], r["kind"], r["fixed"], r["name"]) == (True, "docx", True, RESAVED.name), r
    assert proc.stdout == f"Готово: {copy} — {r['summary']}\n", "строка итога не как у CLI"
    assert data["zip"] is None  # готовый файл один — архив не нужен
    headers, fixed = fetch(srv, r["url"])
    assert headers["Content-Type"] == web.DOCX
    assert f"filename*=UTF-8''{quote(RESAVED.name)}" in headers["Content-Disposition"]
    ours, theirs = members(fixed), members(copy.read_bytes())
    assert ours["word/fontTable.xml"] == theirs["word/fontTable.xml"] != members(original)["word/fontTable.xml"]
    assert ours == theirs, "прочие части отличаются от результата CLI"
    assert RESAVED.read_bytes() == original, "исходник фикстуры изменён"
    assert (srv.root / job / "src" / RESAVED.name).read_bytes() == original, "загруженный файл изменён"
    again = run(srv, job, "landscape")
    assert again["results"] == data["results"] and len(calls) == 1, "ориентация задела .docx"
    assert fetch(srv, again["results"][0]["url"])[1] == fixed


def test_docx_nothing_to_fix_and_not_a_docx(srv, tmp_path):
    """Свежий выход md2docx — «исправлять нечего», скачивать нечего, и имя результату .md рядом
    он не занимает; текст под именем .DOCX — ошибка с тем же текстом, что у CLI."""
    md = tmp_path / "свежий.md"
    md.write_text("# Свежий\n\nТекст и `код`.\n", encoding="utf-8")
    assert cli(md).returncode == 0
    fresh = tmp_path / "свежий.docx"
    fresh_bytes = fresh.read_bytes()
    assert cli("--fix-fonts", fresh).stdout.startswith("Без изменений: ")
    text = tmp_path / "ЗАМЕТКИ.DOCX"
    text.write_text("Просто текст, а не документ Word\n", encoding="utf-8")
    proc = cli("--fix-fonts", text)
    assert proc.returncode == 1
    job = new_job(srv)
    for rel, body in (("свежий.docx", fresh_bytes), ("свежий.md", md.read_bytes()),
                      (text.name, text.read_bytes())):
        assert upload(srv, job, rel, body)[0] == 201
    data = run(srv, job)
    same, converted, bad = data["results"]
    assert (same["ok"], same["kind"], same["fixed"]) == (True, "docx", False), same
    assert not {"url", "name", "summary"} & same.keys()
    assert (converted["ok"], converted["kind"], converted["name"]) == (True, "md", "свежий.docx")
    assert (bad["ok"], bad["kind"], bad["error"]) == (False, "docx", "не docx или файл повреждён")
    assert "url" not in bad and proc.stderr == f"Ошибка: {text}: {bad['error']}\n"
    assert data["zip"] is None  # скачать можно только результат .md
    assert [p for p in stored(srv) if "fonts" in p.parts] == [], "копии без исправлений не удалены"


def test_zip_names_unique_across_md_and_docx(srv):
    """Исправленный «Отчёт.docx» и результаты «отчёт.markdown», «отчёт.md» совпали бы по имени
    (без учёта регистра): второй и третий получают «(2)», «(3)». Имя на странице, при скачивании
    по одному и в архиве — одно и то же, содержимое архива — как у отдельных скачиваний."""
    md = SAMPLE_FILES[0].read_bytes()
    job = new_job(srv)
    for rel, body in (("Отчёт.docx", RESAVED.read_bytes()), ("отчёт.markdown", md), ("отчёт.md", md)):
        assert upload(srv, job, rel, body)[0] == 201
    data = run(srv, job)
    assert [(r["source"], r["kind"], r["name"]) for r in data["results"]] == [
        ("Отчёт.docx", "docx", "Отчёт.docx"), ("отчёт.markdown", "md", "отчёт (2).docx"),
        ("отчёт.md", "md", "отчёт (3).docx")]
    _, archive = fetch(srv, data["zip"])
    with zipfile.ZipFile(io.BytesIO(archive)) as z:
        assert z.testzip() is None
        assert z.namelist() == [r["name"] for r in data["results"]]
        for r in data["results"]:
            headers, single = fetch(srv, r["url"])
            assert f"filename*=UTF-8''{quote(r['name'])}" in headers["Content-Disposition"]
            assert z.read(r["name"]) == single, r["name"]


# --------------------------------------------------------------------------
# Безопасность
# --------------------------------------------------------------------------


@pytest.mark.parametrize("rel", ["../evil.md", "a/../../evil.md", "../../../evil.md", "..", "a/..",
                                 "<abs>", "a\\..\\evil.md", "a\x00evil.md", "", "a//evil.md", "./evil.md"])
def test_unsafe_paths_rejected_400(srv, tmp_path, rel):
    outside = tmp_path / "outside" / "evil.md"
    job = new_job(srv)
    status, _, body = upload(srv, job, rel.replace("<abs>", str(outside)), b"# evil\n")
    assert status == 400, body.decode()
    assert stored(srv) == [], "файл записан"
    assert not outside.exists() and not (tmp_path / "evil.md").exists()


def test_api_requires_token_403(srv):
    assert call(srv, "POST", "/api/jobs", token=False)[0] == 403
    assert call(srv, "POST", "/api/jobs", headers={"X-Token": "not-the-token"})[0] == 403
    assert call(srv, "POST", "/api/jobs?token=not-the-token", token=False)[0] == 403
    job = new_job(srv)
    assert upload(srv, job, "a.md", b"# A\n", token=False)[0] == 403
    assert stored(srv) == []
    assert upload(srv, job, "a.md", b"# A\n")[0] == 201
    assert call(srv, "POST", f"/api/jobs/{job}/convert", b"{}", token=False)[0] == 403
    url = run(srv, job)["results"][0]["url"]
    assert call(srv, "GET", url, token=False)[0] == 403
    assert call(srv, "GET", f"{url}?token={srv.token}", token=False)[0] == 200  # ссылка «Скачать .docx»


def test_foreign_host_403(srv):
    for host in ("evil.example", f"evil.example:{srv.port}", f"127.0.0.1:{srv.port + 1}",
                 f"192.168.0.10:{srv.port}", ""):
        status, _, body = call(srv, "GET", "/", token=False, headers={"Host": host})
        assert status == 403, host
        assert srv.token.encode() not in body
        assert call(srv, "POST", "/api/jobs", headers={"Host": host})[0] == 403, host  # и с токеном
    for host in (f"127.0.0.1:{srv.port}", f"localhost:{srv.port}", f"LOCALHOST:{srv.port}"):
        assert call(srv, "GET", "/", token=False, headers={"Host": host})[0] == 200, host


def test_file_over_limit_413_before_body(srv):
    assert (web.MAX_FILE, web.MAX_JOB) == (60 * 1024 * 1024, 200 * 1024 * 1024)
    job = new_job(srv)
    conn = http.client.HTTPConnection("127.0.0.1", srv.port, timeout=60)
    try:
        conn.putrequest("PUT", f"/api/jobs/{job}/files?path=big.md")
        conn.putheader("X-Token", srv.token)
        conn.putheader("Content-Length", str(web.MAX_FILE + 1))
        conn.endheaders()  # тело не шлём: отказ — по заголовку, до чтения
        resp = conn.getresponse()
        assert resp.status == 413
        assert "60 МБ" in json.loads(resp.read())["error"]
    finally:
        conn.close()
    assert stored(srv) == []


def test_job_over_limit_413(srv):
    srv.max_file, srv.max_job = 1000, 2500
    job = new_job(srv)
    assert upload(srv, job, "a.md", b"a" * 1000)[0] == 201
    assert upload(srv, job, "b.md", b"b" * 1001)[0] == 413  # файл больше лимита
    assert upload(srv, job, "c.md", b"c" * 1000)[0] == 201
    assert upload(srv, job, "d.md", b"d" * 1000)[0] == 413  # вместе 3000 > 2500
    assert upload(srv, job, "e.md", b"e" * 500)[0] == 201   # отказы квоту не съели: ровно 2500
    assert upload(srv, job, "f.md", b"f")[0] == 413
    assert [p.name for p in stored(srv)] == ["a.md", "c.md", "e.md"]
    other = new_job(srv)  # лимит — на задание
    assert upload(srv, other, "g.md", b"g" * 1000)[0] == 201


def test_log_one_line_per_request_without_token_names_or_content(srv, capsys):
    secret = "СОДЕРЖИМОЕ-НЕ-ДЛЯ-ЖУРНАЛА"
    capsys.readouterr()
    job = new_job(srv)
    assert upload(srv, job, "имя-файла.md", f"# {secret}\n".encode())[0] == 201
    url = run(srv, job)["results"][0]["url"]
    assert call(srv, "GET", f"{url}?token={srv.token}", token=False)[0] == 200
    assert call(srv, "POST", "/api/jobs", token=False)[0] == 403
    lines = capsys.readouterr().out.splitlines()
    assert len(lines) == 5, lines
    for line in lines:
        assert LOG_LINE.fullmatch(line), line
    text = "\n".join(lines)
    assert srv.token not in text and secret not in text and "имя-файла" not in text


# --------------------------------------------------------------------------
# Запуск и остановка отдельным процессом
# --------------------------------------------------------------------------


def free_port():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def spawn(cmd, tmpdir, cwd=ROOT):
    """Процесс сервера с TMPDIR в каталоге теста; зависнет — будет убит через минуту."""
    env = dict(os.environ, TMPDIR=str(tmpdir), PYTHONUNBUFFERED="1")
    proc = subprocess.Popen(cmd, cwd=cwd, env=env, text=True,
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    timer = threading.Timer(60, proc.kill)
    timer.start()
    return proc, timer


def banner(proc):
    line = proc.stdout.readline()
    if "http://127.0.0.1:" not in line:
        proc.kill()
        out, err = proc.communicate()
        pytest.fail(f"сервер не стартовал: {line}{out}{err}")
    return line


def finish(proc, timer):
    timer.cancel()
    if proc.poll() is None:
        proc.kill()
    proc.communicate()


def remote(port):
    """Сервер другого процесса: токен — со страницы, как у браузера."""
    probe = SimpleNamespace(port=port, token="")
    page = call(probe, "GET", "/", token=False)[2].decode()
    return SimpleNamespace(port=port, token=re.search(r'const TOKEN = "([^"]+)"', page)[1])


@pytest.mark.parametrize("sig", ["SIGINT", "SIGTERM", "SIGHUP"])
def test_stop_removes_temp_dir(tmp_path, sig):
    tmpdir = tmp_path / "tmp"
    tmpdir.mkdir()
    port = free_port()
    proc, timer = spawn([sys.executable, str(WEB), "--port", str(port), "--no-browser"], tmpdir)
    try:
        assert f"http://127.0.0.1:{port}/" in banner(proc)
        server = remote(port)
        job = new_job(server)
        assert upload(server, job, "a.md", b"# A\n")[0] == 201
        run(server, job)
        (made,) = tmpdir.iterdir()  # один временный каталог процесса: копия и результат
        assert list(made.rglob("a.md")) and list(made.rglob("a.docx"))
        proc.send_signal(getattr(signal, sig))
        assert proc.wait(timeout=30) == 0
    finally:
        finish(proc, timer)
    assert list(tmpdir.iterdir()) == []


def test_busy_port_is_an_error_not_another_port(tmp_path):
    tmpdir = tmp_path / "tmp"
    tmpdir.mkdir()
    with socket.socket() as other:
        other.bind(("127.0.0.1", 0))
        other.listen()
        port = other.getsockname()[1]
        assert web.port_taken(port)
        with pytest.raises(OSError) as exc:
            web.Server(port, tmp_parent=tmpdir)
        assert exc.value.errno == web.errno.EADDRINUSE
        proc = subprocess.run([sys.executable, str(WEB), "--port", str(port), "--no-browser"],
                              env=dict(os.environ, TMPDIR=str(tmpdir)),
                              capture_output=True, text=True, timeout=60)
    assert proc.returncode == 1, proc.stdout + proc.stderr
    assert f"порт {port} занят" in proc.stderr
    assert proc.stdout == "", "сервер запустился на другом порту"
    assert list(tmpdir.iterdir()) == []
    assert not web.port_taken(port)


def test_command_file_starts_server_from_its_folder(tmp_path):
    cmd = ROOT / "md2docx-web.command"
    assert os.access(cmd, os.X_OK), "md2docx-web.command без права на запуск"
    tmpdir = tmp_path / "tmp"
    tmpdir.mkdir()
    port = free_port()
    proc, timer = spawn([str(cmd), "--port", str(port), "--no-browser"], tmpdir, cwd=tmp_path)
    try:
        line = banner(proc)
        assert f"md2docx v{web.md2docx.__version__}" in line and f"http://127.0.0.1:{port}/" in line
        assert call(SimpleNamespace(port=port, token=""), "GET", "/", token=False)[0] == 200
        proc.send_signal(signal.SIGTERM)
        assert proc.wait(timeout=30) == 0
    finally:
        finish(proc, timer)
    assert list(tmpdir.iterdir()) == []


def test_readme_section_names_the_pinned_port():
    for name, title in (("README.md", "## Browser interface"), ("README.ru.md", "## Интерфейс в браузере")):
        text = (ROOT / name).read_text(encoding="utf-8")
        assert text.count(title) == 1, name
        section = text.split(title, 1)[1].split("\n## ", 1)[0]
        assert str(web.PORT) in section and "md2docx-web.command" in section, name
        assert str(web.PORT + 10000) in section, f"{name}: не назван парный порт (frontend + 10000)"
    assert 51000 <= web.PORT <= 53000
