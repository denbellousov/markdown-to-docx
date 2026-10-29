"""Перевод интерфейса: каталоги locales/*.json полны и согласованы с кодом, язык выбирается сам —
MD2DOCX_LANG, переменные локали, язык ОС, иначе английский; в браузере — заголовок X-Lang.

    python3 -m pytest tests/test_i18n.py -q --basetemp=<свой каталог> -p no:cacheprovider
"""

from __future__ import annotations

import ast
import http.client
import json
import os
import re
import string
import subprocess
import sys
import threading
import zipfile
from pathlib import Path, PurePosixPath

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import md2docx  # noqa: E402
import md2docx_web as web  # noqa: E402

MODULES = ("md2docx.py", "md2docx_web.py", "md2docx_app.py")
LANGS = sorted(set(md2docx.LANGUAGES) - {"en"})
ENV = ("MD2DOCX_LANG", "LANGUAGE", "LC_ALL", "LC_MESSAGES", "LANG")
DYNAMIC = {"en": md2docx.Parser.MESSAGES}  # tr(en) в Parser.error перебирает эти строки
RESAVED = ROOT / "tests" / "fixtures" / "02-tables-resaved.docx"
SUMMARY = "sections: {sections} ({letters}), tables: {tables}, auto-landscape: {auto}"


def code_keys():
    """Строки, которые код передаёт в tr() → (ключи, вызовы, ключ которых не вычислить):
    литералы, строковые константы модуля и DYNAMIC."""
    keys, unresolved = set(), []
    for name in MODULES:
        tree = ast.parse((ROOT / name).read_text(encoding="utf-8"))
        consts = {target.id: node.value.value for node in tree.body if isinstance(node, ast.Assign)
                  for target in node.targets if isinstance(target, ast.Name)
                  and isinstance(node.value, ast.Constant) and isinstance(node.value.value, str)}
        for node in ast.walk(tree):
            func = getattr(node, "func", None)
            if not isinstance(node, ast.Call) or getattr(func, "id", getattr(func, "attr", None)) != "tr":
                continue
            arg = node.args[0]
            if isinstance(arg, ast.Constant) and isinstance(arg.value, str):
                keys.add(arg.value)
            elif isinstance(arg, ast.Name) and arg.id in consts:
                keys.add(consts[arg.id])
            elif isinstance(arg, ast.Name) and arg.id in DYNAMIC:
                keys.update(DYNAMIC[arg.id])
            else:
                unresolved.append(f"{name}:{node.lineno}")
    return keys, unresolved


def catalog(lang):
    return json.loads((ROOT / "locales" / f"{lang}.json").read_text(encoding="utf-8"))


def fields(text):
    return {name for _, name, _, _ in string.Formatter().parse(text) if name is not None}


def env_with(lang):
    return {k: v for k, v in os.environ.items() if k not in ENV} | {"MD2DOCX_LANG": lang}


# --------------------------------------------------------------------------
# Каталоги
# --------------------------------------------------------------------------


def test_every_tr_call_is_checkable():
    keys, unresolved = code_keys()
    assert unresolved == [], "ключ перевода не виден тесту"
    assert "Done: {file} — {summary}" in keys and md2docx.NOT_TEXT in keys and web.GONE in keys


def test_one_catalog_per_language():
    assert sorted(p.stem for p in (ROOT / "locales").glob("*.json")) == LANGS
    assert md2docx.LANGUAGES["en"] == "English" and len(md2docx.LANGUAGES) == 10


@pytest.mark.parametrize("lang", LANGS)
def test_catalog_complete_and_consistent(lang):
    keys = code_keys()[0] | web.PAGE_KEYS
    cat = catalog(lang)
    assert sorted(keys - cat.keys()) == [], "нет перевода"
    assert sorted(cat.keys() - keys) == [], "перевод строки, которой в коде нет"
    assert [k for k, v in cat.items() if not v.strip() or fields(v) != fields(k)] == [], \
        "пустой перевод или другие {поля}"
    assert len(cat["P"]) <= 2 and len(cat["L"]) <= 2 and cat["P"] != cat["L"], "буквы разделов"


def test_page_keys_cover_markup_and_script():
    """Ключи страницы собраны так же, как их собирает её скрипт: вложенный элемент — {0}."""
    assert "Same as {0}." in web.i18n_key("Same as <code>--landscape</code>.")
    assert {"Choose files", "Could not convert: {error}", "Results",
            "The whole document is landscape; markers in the text still switch orientation. Same as {0}.",
            } <= web.PAGE_KEYS


# --------------------------------------------------------------------------
# Выбор языка
# --------------------------------------------------------------------------


@pytest.mark.parametrize("tags, expected", [
    (["pt_BR.UTF-8"], "pt"), (["zh-Hans-CN"], "zh"), (["sr@latin", "de-AT"], "de"),
    (["JA"], "ja"), (["sv", "nb"], "en"), ([], "en"),
])
def test_pick_language(tags, expected):
    assert md2docx.pick_language(tags) == expected


@pytest.mark.parametrize("env, expected", [
    ({"MD2DOCX_LANG": "de", "LC_ALL": "fr_FR.UTF-8"}, "de"),
    ({"LANGUAGE": "xx:pt_BR", "LANG": "fr_FR.UTF-8"}, "pt"),
    ({"LC_ALL": "es_ES.UTF-8", "LANG": "fr_FR.UTF-8"}, "es"),
    ({"LC_MESSAGES": "hi_IN.UTF-8", "LANG": "fr_FR.UTF-8"}, "hi"),
    ({"LANG": "zh_CN.UTF-8"}, "zh"),
    ({"LANG": "C.UTF-8"}, "ja"),        # C и POSIX — не язык: берётся язык ОС
    ({"LC_ALL": "POSIX"}, "ja"),
    ({"LANG": "sv_SE.UTF-8"}, "en"),    # язык задан, но перевода нет — английский
    ({}, "ja"),
])
def test_system_language(monkeypatch, env, expected):
    for name in ENV:
        monkeypatch.delenv(name, raising=False)
    for name, value in env.items():
        monkeypatch.setenv(name, value)
    monkeypatch.setattr(md2docx, "os_languages", lambda: ["ja-JP", "en-US"])
    md2docx.system_language.cache_clear()
    try:
        assert md2docx.system_language() == expected
    finally:
        md2docx.system_language.cache_clear()


def test_english_needs_no_catalog():
    token = md2docx.LANG.set("en")
    try:
        assert md2docx.tr("Done: {file} — {summary}", file="a.docx", summary="x") == "Done: a.docx — x"
        assert md2docx.tr("P") == "P" and md2docx.language() == "en"
    finally:
        md2docx.LANG.reset(token)


def test_right_to_left_isolates_names_only_for_the_page(tmp_path):
    """Справа налево имя «02-a.md» без изоляции видно как «a.md-02». Сообщения страницы (язык
    задан на запрос) получают строки и пути в FSI…PDI; терминал и сам документ — без невидимых
    символов: они попали бы в скопированный путь и в текст. В ar.json расширения файлов — после
    LRM, по той же причине."""
    src = tmp_path / "02-a.md"
    src.write_text("![x](missing.png)\n", encoding="utf-8")
    token = md2docx.LANG.set("ar")
    try:
        for file in ("02-a.md", PurePosixPath("02-a.md")):  # путь загрузки — PurePosixPath
            assert md2docx.tr("{file}: could not save", file=file).startswith("⁨02-a.md⁩"), repr(file)
        stats = md2docx.convert(src, tmp_path / "page.docx")
    finally:
        md2docx.LANG.reset(token)
    assert re.search("[⁦-⁩]", stats["warnings"][0])
    with zipfile.ZipFile(tmp_path / "page.docx") as z:
        body = z.read("word/document.xml").decode()
    text = "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", body))  # заглушка — в нескольких прогонах
    assert catalog("ar")["[image: {alt}]"].format(alt="x") in text
    assert not re.search("[⁦-⁩]", body), "изолирующие символы страницы попали в документ"
    proc = subprocess.run([sys.executable, str(ROOT / "md2docx.py"), str(src)], env=env_with("ar"),
                          capture_output=True, text=True, timeout=120)
    assert proc.returncode == 0 and "02-a.md" in proc.stderr, proc.stderr
    assert not re.search("[⁦-⁩]", proc.stdout + proc.stderr)
    bare = {k: v for k, v in catalog("ar").items() if re.search(r"(?<!‎)\.(md|markdown|docx|zip)\b", v)}
    assert bare == {}, "расширение без LRM перевернётся: .docx → docx."


def test_python_api_letters_do_not_depend_on_language(tmp_path):
    """convert()["sections"] — К и А, как в 1.0.0, на любом языке; переводится только строка итога."""
    src = tmp_path / "a.md"
    src.write_text("# A\n\n<!-- landscape -->\n\nB\n", encoding="utf-8")
    token = md2docx.LANG.set("en")
    try:
        stats = md2docx.convert(src, tmp_path / "a.docx")
        assert stats["sections"] == ["К", "А"]
        assert md2docx.summary(stats).startswith("sections: 2 (P, L), ")
    finally:
        md2docx.LANG.reset(token)


@pytest.mark.parametrize("lang", ["en", *LANGS])
def test_verify_accepts_the_placeholder_in_every_language(tmp_path, lang):
    """tests/verify.py сверяет документ на любом языке интерфейса: заглушка картинки — из каталогов."""
    src, out = tmp_path / "a.md", tmp_path / "a.docx"
    src.write_text("# A\n\nText ![company logo](missing.png) text.\n", encoding="utf-8")
    token = md2docx.LANG.set(lang)
    try:
        md2docx.convert(src, out)
    finally:
        md2docx.LANG.reset(token)
    proc = subprocess.run([sys.executable, str(ROOT / "tests" / "verify.py"), str(src), str(out)],
                          capture_output=True, text=True, timeout=120)
    assert proc.returncode == 0, proc.stdout + proc.stderr


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------


def test_cli_in_english(tmp_path):
    src = tmp_path / "a.md"
    src.write_text("# A\n\n![logo](missing.png)\n\n| a | b |\n|---|---|\n| 1 | 2 |\n", encoding="utf-8")
    proc = subprocess.run([sys.executable, str(ROOT / "md2docx.py"), str(src)], env=env_with("en"),
                          capture_output=True, text=True, timeout=120)
    assert proc.returncode == 0, proc.stderr
    assert proc.stdout == f"Done: {tmp_path / 'a.docx'} — sections: 1 (P), tables: 1, auto-landscape: 0\n"
    assert proc.stderr == ("Warning: a.md, line 3: image not found or remote: missing.png"
                           " — a placeholder is inserted\n")


@pytest.mark.parametrize("lang", LANGS)
def test_cli_help_and_errors_translated(lang):
    cat = catalog(lang)
    cmd = [sys.executable, str(ROOT / "md2docx.py")]
    shown = subprocess.run([*cmd, "--help"], env=env_with(lang), capture_output=True, text=True, timeout=60)
    assert shown.returncode == 0 and shown.stdout.startswith(cat["usage:"] + " md2docx "), shown.stdout
    flat = " ".join(shown.stdout.split())
    for key in ("show the version and exit", "positional arguments", "options"):
        assert " ".join(cat[key].split()) in flat, key
    failed = subprocess.run(cmd, env=env_with(lang), capture_output=True, text=True, timeout=60)
    message = cat["the following arguments are required"] + ": INPUT.md"
    assert failed.returncode == 2
    assert failed.stderr.splitlines()[-1] == cat["{prog}: error: {message}"].format(prog="md2docx", message=message)


# --------------------------------------------------------------------------
# Браузер: язык страницы приходит в X-Lang
# --------------------------------------------------------------------------


@pytest.fixture
def srv(tmp_path):
    server = web.Server(0, tmp_parent=tmp_path)
    threading.Thread(target=server.serve_forever, kwargs={"poll_interval": 0.05}, daemon=True).start()
    yield server
    server.shutdown()
    server.server_close()
    server.cleanup()


def call(srv, method, path, body=None, lang=None):
    conn = http.client.HTTPConnection("127.0.0.1", srv.port, timeout=120)
    headers = {"X-Token": srv.token} | ({"X-Lang": lang} if lang is not None else {})
    try:
        conn.request(method, path, body=body, headers=headers)
        resp = conn.getresponse()
        return resp.status, resp.read()
    finally:
        conn.close()


def test_server_answers_in_the_page_language(srv, monkeypatch):
    """Ошибки и итоги — на языке из X-Lang; неизвестный язык или его отсутствие — язык системы
    (в тестах русский). .docx исправляется один раз: другой язык меняет только текст — итога
    и ошибки. Ошибка хранится без трассировки: та держала бы распакованный docx до нового задания."""
    for lang, expected in (("en", "no such address"), ("de", catalog("de")["no such address"]),
                           ("xx", catalog("ru")["no such address"]), (None, catalog("ru")["no such address"])):
        status, body = call(srv, "GET", "/api/nope", lang=lang)
        assert (status, json.loads(body)) == (404, {"error": expected}), lang
    calls, real = [], md2docx.fix_fonts
    monkeypatch.setattr(md2docx, "fix_fonts", lambda path, **kw: calls.append(path) or real(path, **kw))
    job = json.loads(call(srv, "POST", "/api/jobs")[1])["job"]
    assert call(srv, "PUT", f"/api/jobs/{job}/files?path=a.md", b"# A\n")[0] == 201
    assert call(srv, "PUT", f"/api/jobs/{job}/files?path={RESAVED.name}", RESAVED.read_bytes())[0] == 201
    assert call(srv, "PUT", f"/api/jobs/{job}/files?path=broken.docx", b"not a zip")[0] == 201
    results = {}
    for lang in ("en", "de", "ar"):
        status, body = call(srv, "POST", f"/api/jobs/{job}/convert", b'{"orientation": "auto"}', lang=lang)
        assert status == 200, body
        results[lang] = json.loads(body)["results"]
    assert results["en"][0]["summary"] == "sections: 1 (P), tables: 0, auto-landscape: 0"
    fonts = results["en"][1]["summary"].split(": ", 1)[1]
    for lang, iso in (("de", "{}"), ("ar", "⁨{}⁩")):  # справа налево строки изолированы, числа — нет
        cat = catalog(lang)
        assert results[lang][0]["summary"] == cat[SUMMARY].format(sections=1, letters=iso.format(cat["P"]),
                                                                  tables=0, auto=0)
        assert results[lang][1]["summary"] == cat["font hints restored: {fonts}"].format(fonts=iso.format(fonts))
    assert results["en"][1]["summary"].startswith("font hints restored: ")
    broken = "not a docx file or the file is damaged"
    assert {lang: results[lang][2]["error"] for lang in results} == \
        {"en": broken, "de": catalog("de")[broken], "ar": catalog("ar")[broken]}
    assert len(calls) == 2, "смена языка заново исправляла .docx"
    (exc,) = [exc for _, exc, _ in srv.jobs[job].fonts.values() if exc]
    assert (exc.__traceback__, exc.__cause__, exc.__context__) == (None, None, None)


def test_page_carries_every_translation(srv):
    status, body = call(srv, "GET", "/")
    page = body.decode()
    assert status == 200 and '<html lang="en" data-theme="dark">' in page
    i18n = json.loads(re.search(r"const I18N = (.*);\n", page)[1])
    assert sorted(i18n) == LANGS
    for lang in LANGS:
        assert i18n[lang] == {k: v for k, v in catalog(lang).items() if k in web.PAGE_KEYS}, lang
        assert i18n[lang].keys() == web.PAGE_KEYS, lang
    assert json.loads(re.search(r"const LANGS = (.*);\n", page)[1]) == [list(x) for x in md2docx.LANGUAGES.items()]
    assert "const APP = false;" in page
    assert web.script_json("</script>") == '"\\u003c/script>"'
