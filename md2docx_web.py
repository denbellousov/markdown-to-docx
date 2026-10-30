#!/usr/bin/env python3
"""md2docx_web — простой локальный интерфейс md2docx в браузере.

    python3 md2docx_web.py [--port N] [--no-browser]

Печатает адрес и открывает браузер; Ctrl+C останавливает. Сервер слушает только 127.0.0.1,
API отвечает только с токеном этого запуска. .md конвертируются, в .docx восстанавливаются
подсказки шрифтов (как --fix-fonts). Файлы из браузера — копии: они лежат во временном каталоге
процесса, хранится только последнее задание, при выходе удаляется всё; исходники на диске
не меняются. Только стандартная библиотека и md2docx из этого же каталога.
"""

from __future__ import annotations

import argparse
import atexit
import errno
import functools
import html
import json
import os
import re
import secrets
import shutil
import signal
import socket
import sys
import tempfile
import threading
import time
import webbrowser
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path, PurePosixPath, PureWindowsPath
from urllib.parse import parse_qs, quote, urlsplit

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))  # md2docx — из этого каталога, а не установленный где-то ещё
try:
    import md2docx  # noqa: E402
except ModuleNotFoundError as exc:  # pragma: no cover — на машине без зависимостей; tr() ещё нет
    sys.exit(f"Error: module {exc.name} not found — run: pip install -r requirements.txt")
from md2docx import tr  # noqa: E402

PORT = 51536            # закреплён за проектом (парный 61536 не используется); другой — только --port
HOST = "127.0.0.1"
MB = 1024 * 1024
MAX_FILE = 60 * MB      # на один файл
MAX_JOB = 200 * MB      # на задание — все файлы одного выбора в браузере
MAX_JSON = 64 * 1024    # тело запроса конвертации
CHUNK = MB
KEEPALIVE = 10          # секунд между пробелами в ответе, пока открыто окно «Сохранить» приложения
MD_SUFFIXES = (".md", ".markdown")
DOCX_SUFFIX = ".docx"   # готовый документ: вернуть подсказки шрифтов, как --fix-fonts
GONE = "the results are out of date — choose the files again"  # задание удалено: его заменило новое
MODES = {  # переключатель на странице → параметры convert(), как флаги CLI
    "auto": ("Auto", {}),
    "landscape": ("All landscape", {"landscape": True}),
    "no-auto": ("No auto-landscape", {"auto_landscape": False}),
}
DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


# ---------------------------------------------------------------- помощники

def safe_rel(raw):
    """Относительный путь из браузера → PurePosixPath; None — если он может выйти из задания."""
    if not raw or "\x00" in raw or "\\" in raw or raw.startswith("/"):
        return None
    parts = raw.split("/")
    if any(part in ("", ".", "..") or PureWindowsPath(part).drive for part in parts):  # D: — другой диск Windows
        return None
    return PurePosixPath(*parts)


def port_taken(port):
    """Порт уже слушают на 127.0.0.1 — отдельно или на всех адресах сразу.

    Проверка нужна до bind: с SO_REUSEADDR (без него не перезапуститься, пока старые
    соединения в TIME_WAIT) macOS молча даёт занять 127.0.0.1:порт поверх чужого *:порт."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(1)
        return s.connect_ex((HOST, port)) == 0


def stamp():
    """Время строки журнала — с часовым поясом (в Москве GMT+3)."""
    t = time.localtime()
    off = t.tm_gmtoff // 60
    zone = f"GMT{'+' if off >= 0 else '-'}{abs(off) // 60}" + (f":{abs(off) % 60:02d}" if off % 60 else "")
    return time.strftime("%H:%M:%S ", t) + zone


def disposition(name):
    """Content-Disposition: запасное ASCII-имя и полное имя в UTF-8 (RFC 6266)."""
    fallback = re.sub(r"[^A-Za-z0-9._() -]", "_", name)
    return f"attachment; filename=\"{fallback}\"; filename*=UTF-8''{quote(name, safe='')}"


def result_name(rel, taken):
    """Путь результата в задании → PurePosixPath «…/имя.docx»; если он уже в taken (без учёта
    регистра) — «имя (2).docx», «имя (3).docx»… Под этим именем файл и в архиве, и скачивается."""
    out, n = rel.with_suffix(".docx"), 1
    while str(out).casefold() in taken:  # a.md и a.markdown, a.docx и a.md: иначе второй затёр бы первый
        n += 1
        out = rel.with_name(f"{rel.stem} ({n}).docx")
    taken.add(str(out).casefold())
    return out


def hide_paths(text, job):
    """Пути временного каталога в тексте ошибки → относительные: браузеру они ни к чему."""
    for base in (job.src, job.dir):  # /private/var/… раньше /var/…: второй — часть первого
        for variant in sorted({str(base), os.path.realpath(base)}, key=len, reverse=True):
            text = text.replace(variant + os.sep, "").replace(variant, "")
    return text


# ---------------------------------------------------------------- задания

class Job:
    """Файлы одного выбора в браузере и результаты последнего прогона."""

    def __init__(self, job_id, path):
        self.id = job_id
        self.dir = path
        self.src = path / "src"
        self.src.mkdir(parents=True)
        self.files = []       # загруженные пути (PurePosixPath) в порядке загрузки
        self.size = 0         # байт загружено и зарезервировано под идущие загрузки
        self.runs = 0
        self.run_dir = None   # каталог результатов последнего прогона
        self.results = []
        self.fonts = {}       # .docx → итог исправления шрифтов: от ориентации не зависит


def fix_copy(job, rel):
    """Вернуть подсказки шрифтов в копию .docx задания → (исправленное, ошибка, путь копии).

    Загруженный файл в src не меняется: исправляется его копия в fonts/. Если менять нечего
    или ошибка, копия удаляется — скачивать нечего."""
    src, dest = job.src.joinpath(*rel.parts), job.dir.joinpath("fonts", *rel.parts)
    try:
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(src, dest)
        changed = md2docx.fix_fonts(dest, force=True)  # force: эту копию в Word никто не открывал
    except Exception as exc:  # текст ошибки — в fonts_item, на языке запроса
        dest.unlink(missing_ok=True)
        exc.__traceback__ = exc.__cause__ = exc.__context__ = None  # кадры держали бы распакованный docx
        return None, exc, dest
    if not changed:
        dest.unlink(missing_ok=True)
    return changed, None, dest


def fonts_item(job, rel, changed, exc, dest):
    """Итог fix_copy → результат для списка на языке текущего запроса."""
    item = {"source": str(rel), "kind": "docx", "warnings": []}
    if isinstance(exc, (OSError, ValueError)):
        item.update(ok=False, error=hide_paths(md2docx.describe_error(exc, dest), job))
    elif exc:  # непредвиденное — в ответ, как «внутренняя ошибка» в CLI
        item.update(ok=False, error=hide_paths(md2docx.internal_error(exc), job))
    else:
        item.update(ok=True, fixed=bool(changed))
        if changed:
            item.update(summary=md2docx.fonts_summary(changed), file=dest)
    return item


def run_job(job, sources, options):
    """Обработать файлы задания → (каталог прогона, результаты в порядке загрузки): .md —
    сконвертировать в новый каталог прогона, в .docx — исправить шрифты (один раз на задание:
    ориентация и язык их не касаются, повторный прогон берёт готовый итог).

    Как пакетный запуск CLI: ошибка в одном файле не останавливает остальные."""
    job.runs += 1
    run_dir = job.dir / f"run{job.runs}"
    taken, results = set(), []
    for rel in sources:
        if rel.suffix.lower() == DOCX_SUFFIX:
            if rel not in job.fonts:
                job.fonts[rel] = fix_copy(job, rel)
            item = fonts_item(job, rel, *job.fonts[rel])
            if "file" in item:  # имя занимает только то, что можно скачать
                out_rel = result_name(rel, taken)
                item.update(name=out_rel.name, zipname=str(out_rel))
            results.append(item)
            continue
        out_rel = result_name(rel, taken)
        src, dest = job.src.joinpath(*rel.parts), run_dir.joinpath(*out_rel.parts)
        item = {"source": str(rel), "kind": "md", "name": out_rel.name, "warnings": []}
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            stats = md2docx.convert(src, dest, **options)
        except (OSError, UnicodeDecodeError) as exc:
            item.update(ok=False, error=hide_paths(md2docx.describe_error(exc, src), job))
        except Exception as exc:  # непредвиденное — в ответ, как «внутренняя ошибка» в CLI
            item.update(ok=False, error=hide_paths(md2docx.internal_error(exc), job))
        else:
            item.update(ok=True, warnings=stats["warnings"], summary=md2docx.summary(stats),
                        file=dest, zipname=str(out_rel))
        results.append(item)
    return run_dir, results


# ---------------------------------------------------------------- сервер

class Server(ThreadingHTTPServer):
    daemon_threads = True

    def __init__(self, port=PORT, tmp_parent=None, app=False, save_as=None):
        if port and port_taken(port):
            raise OSError(errno.EADDRINUSE, os.strerror(errno.EADDRINUSE))
        super().__init__((HOST, port), Handler)
        self.app = app  # запущен приложением md2docx, а не из терминала: так и подсказывает страница
        self.save_as = save_as  # приложение: окно «Сохранить» по имени файла → путь или None («Отмена»)
        self.port = self.server_address[1]
        self.hosts = {f"127.0.0.1:{self.port}", f"localhost:{self.port}"}
        self.token = secrets.token_urlsafe(32)
        self.root = Path(tempfile.mkdtemp(prefix="md2docx-web-", dir=tmp_parent))
        self.jobs = {}
        self.lock = threading.Lock()          # состояние заданий
        self.convert_lock = threading.Lock()  # прогоны и сборка zip — по одному
        self.max_file, self.max_job = MAX_FILE, MAX_JOB

    def cleanup(self):
        shutil.rmtree(self.root, ignore_errors=True)


class Handler(BaseHTTPRequestHandler):
    server_version = "md2docx-web"
    sys_version = ""
    timeout = 120  # зависшее соединение не держит поток вечно

    # --- журнал: одна строка на запрос; без query (там токен и имена файлов) и без содержимого

    def log_request(self, code="-", size="-"):
        path = urlsplit(getattr(self, "path", "") or "").path
        sys.stdout.write(f"{stamp()}  {self.command or '-'} {path} {getattr(code, 'value', code)}\n")
        sys.stdout.flush()

    def log_message(self, format, *args):  # noqa: A002 — строка запроса уже в log_request
        pass

    # --- разбор запроса

    def do_GET(self):
        try:
            self.dispatch()
        except (ConnectionError, TimeoutError):  # браузер отменил скачивание, клиент ушёл
            self.close_connection = True

    do_POST = do_PUT = do_GET

    def dispatch(self):
        lang = self.headers.get("X-Lang", "")  # язык, выбранный на странице; иначе — системный
        md2docx.LANG.set(lang if lang in md2docx.LANGUAGES else None)
        length = self.headers.get("Content-Length", "0")
        valid = re.fullmatch(r"[0-9]{1,15}", length) and "Transfer-Encoding" not in self.headers
        self.body_left = int(length) if valid else 0
        url = urlsplit(self.path)
        self.query = parse_qs(url.query, keep_blank_values=True)
        if self.headers.get("Host", "").lower() not in self.server.hosts:
            return self.fail(403, tr("access only via 127.0.0.1 or localhost"))
        if not valid:
            return self.fail(411, tr("the Content-Length header is required"))
        if url.path == "/" and self.command == "GET":
            return self.page()
        parts = url.path.split("/")[1:]
        if parts[:1] != ["api"]:
            return self.fail(404, tr("no such page"))
        token = self.headers.get("X-Token") or self.query.get("token", [""])[0]
        if not secrets.compare_digest(token.encode(), self.server.token.encode()):
            return self.fail(403, tr("no token, or it is from a previous run — reload the page"))
        return self.route(parts[1:])

    def route(self, parts):
        method = self.command
        if parts == ["jobs"] and method == "POST":
            return self.create_job()
        if len(parts) >= 3 and parts[0] == "jobs":
            with self.server.lock:
                job = self.server.jobs.get(parts[1])
            if job is None:
                return self.fail(404, tr(GONE))
            action = parts[2:]
            if action == ["files"] and method == "PUT":
                return self.upload(job)
            if action == ["convert"] and method == "POST":
                return self.convert(job)
            take = method == "GET" or (method == "POST" and self.server.save_as is not None)  # POST — сохранить
            if action == ["zip"] and take:
                return self.zip(job)
            if (len(action) == 2 and action[0] == "docx" and take
                    and re.fullmatch(r"[0-9]{1,9}", action[1])):
                return self.download(job, int(action[1]))
        return self.fail(404, tr("no such address"))

    # --- ответы

    def common_headers(self):
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")

    def send(self, code, body, ctype, headers=()):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.common_headers()
        for name, value in headers:
            self.send_header(name, value)
        self.end_headers()
        self.wfile.write(body)

    def reply(self, code, obj):
        self.send(code, json.dumps(obj, ensure_ascii=False).encode(), "application/json; charset=utf-8")

    def fail(self, code, message):
        # непрочитанное тело дочитываем (если оно не больше лимита файла): закрытие сокета
        # с данными в буфере даёт RST, и клиент видит обрыв соединения вместо ответа
        if self.body_left <= self.server.max_file:
            self.discard()
        self.reply(code, {"error": message})

    def discard(self):
        try:
            while self.body_left > 0:
                chunk = self.rfile.read(min(CHUNK, self.body_left))
                if not chunk:
                    break
                self.body_left -= len(chunk)
        except OSError:  # клиент ушёл или замолчал — отвечать уже некому
            self.body_left = 0

    def send_file(self, path, ctype, name):
        try:
            f = open(path, "rb")
        except FileNotFoundError:  # прогон успели заменить новым
            return self.fail(404, tr("the result has been updated — download it again"))
        with f:
            self.send_response(200)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(os.fstat(f.fileno()).st_size))
            self.send_header("Content-Disposition", disposition(name))
            self.common_headers()
            self.end_headers()
            shutil.copyfileobj(f, self.wfile, CHUNK)

    def deliver(self, path, ctype, name):
        """GET — файл в ответе. POST (приложение) — файл туда, куда укажут в окне «Сохранить»:
        загрузка WebKit поверх существующего файла молча не пишет, а «Заменить» в окне уже нажали."""
        if self.command == "GET":
            return self.send_file(path, ctype, name)
        try:
            f = open(path, "rb")  # до окна: пока оно открыто, прогон могут заменить новым
        except FileNotFoundError:
            return self.fail(404, tr("the result has been updated — download it again"))
        # окно открыто сколько угодно, а WebKit ждёт данных не больше минуты: заголовки — сразу,
        # пока окно открыто — пробел раз в KEEPALIVE секунд (JSON их пропускает), итог — в теле
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.common_headers()
        self.end_headers()
        done = threading.Event()

        def keepalive():
            try:
                while not done.wait(KEEPALIVE):
                    self.wfile.write(b" ")
            except OSError:  # окно закрыли — ответ уже некому читать
                pass

        writer = threading.Thread(target=keepalive, daemon=True)
        writer.start()
        with f:
            try:
                target = self.server.save_as(name)
                if target:
                    with open(target, "wb") as out:
                        shutil.copyfileobj(f, out, CHUNK)
                answer = {"saved": bool(target)}
            except OSError as exc:
                answer = {"error": md2docx.describe_error(exc, None)}
        done.set()
        writer.join()
        self.wfile.write(json.dumps(answer, ensure_ascii=False).encode())

    # --- действия

    def page(self):
        nonce = secrets.token_urlsafe(16)
        values = {"NONCE": nonce, "TOKEN": self.server.token, "VERSION": html.escape(md2docx.__version__),
                  "MAX_FILE": str(self.server.max_file), "MAX_JOB": str(self.server.max_job),
                  "I18N": page_catalogs(), "LANGS": script_json(list(md2docx.LANGUAGES.items())),
                  "APP": script_json(self.server.app),
                  # окно приложения называет язык своего пакета, а не выбранный в системе
                  "SYSLANG": script_json(md2docx.system_language() if self.server.app else None)}
        body = re.sub(r"__(NONCE|TOKEN|VERSION|MAX_FILE|MAX_JOB|I18N|LANGS|APP|SYSLANG)__",
                      lambda m: values[m[1]], PAGE)
        csp =(f"default-src 'none'; script-src 'nonce-{nonce}'; style-src 'unsafe-inline'; "
               "img-src data:; connect-src 'self'; form-action 'none'; frame-ancestors 'none'; "
               "base-uri 'none'")
        self.send(200, body.encode(), "text/html; charset=utf-8", [("Content-Security-Policy", csp)])

    def create_job(self):
        if self.body_left > MAX_JSON:
            return self.fail(413, tr("the request is too large"))
        self.discard()
        job_id = secrets.token_hex(8)
        # страница показывает только последнее задание: прежние удаляются с диска целиком;
        # под convert_lock — чтобы не посреди их прогона или сборки архива
        with self.server.convert_lock:
            job = Job(job_id, self.server.root / job_id)
            with self.server.lock:
                self.server.jobs = {job_id: job}
            for path in self.server.root.iterdir():
                if path != job.dir:
                    shutil.rmtree(path, ignore_errors=True)
        self.reply(201, {"job": job_id})

    def current(self, job):
        """Задание всё ещё последнее — его не удалило новое (create_job)."""
        with self.server.lock:
            return self.server.jobs.get(job.id) is job

    def upload(self, job):
        rel = safe_rel(self.query.get("path", [""])[0])
        if rel is None:
            return self.fail(400, tr("invalid file path: it must be relative, without “..”, “\\” or NUL"))
        if "Content-Length" not in self.headers:
            return self.fail(411, tr("the Content-Length header is required"))
        size = self.body_left
        if size > self.server.max_file:
            return self.fail(413, tr("{file}: the file is larger than {size} MB", file=rel,
                                     size=self.server.max_file // MB))
        with self.server.lock:
            fits = job.size + size <= self.server.max_job
            if fits:
                job.size += size
        if not fits:
            return self.fail(413, tr("the files together exceed {size} MB — choose fewer files",
                                     size=self.server.max_job // MB))
        dest = job.src.joinpath(*rel.parts)  # без «..» и «/» в начале — внутри каталога задания
        created = written = False
        error = (500, tr("{file}: could not save", file=rel))
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            with open(dest, "xb") as out:  # «x»: занятое имя — отказ, а не перезапись
                created = True
                while self.body_left:
                    chunk = self.rfile.read(min(CHUNK, self.body_left))
                    if not chunk:
                        raise ConnectionError
                    out.write(chunk)
                    self.body_left -= len(chunk)
            written = True
        except (FileExistsError, NotADirectoryError, IsADirectoryError):
            error = (409, tr("{file}: a file with this path is already uploaded", file=rel))
        except (ConnectionError, TimeoutError):
            error = (400, tr("{file}: the file was not transferred completely", file=rel))
        except OSError as exc:
            error = (500, tr("{file}: could not save ({error})", file=rel,
                             error=exc.strerror or exc.__class__.__name__))
        finally:
            if not written:
                with self.server.lock:
                    job.size -= size
                if created:
                    dest.unlink(missing_ok=True)
        if not written:
            return self.fail(*error)
        with self.server.lock:
            job.files.append(rel)
        self.reply(201, {"path": str(rel), "size": size})

    def convert(self, job):
        if self.body_left > MAX_JSON:
            return self.fail(413, tr("the request is too large"))
        raw = self.rfile.read(self.body_left) if self.body_left else b"{}"
        self.body_left = 0
        try:
            data = json.loads(raw)
        except ValueError:  # в том числе не UTF-8
            data = None
        mode = data.get("orientation", "auto") if isinstance(data, dict) else None
        if mode not in MODES:
            return self.fail(400, tr("orientation must be auto, landscape or no-auto"))
        with self.server.lock:
            sources = [rel for rel in job.files if rel.suffix.lower() in (*MD_SUFFIXES, DOCX_SUFFIX)]
        if not sources:
            return self.fail(400, tr("there are no .md or .docx files among the uploaded ones"))
        with self.server.convert_lock:  # convert() не рассчитан на параллельные вызовы
            gone = not self.current(job)  # пока запрос ждал очереди, выбрали новые файлы
            if not gone:
                run_dir, results = run_job(job, sources, MODES[mode][1])
                with self.server.lock:
                    old, job.run_dir, job.results = job.run_dir, run_dir, results
                if old:
                    shutil.rmtree(old, ignore_errors=True)
        if gone:
            return self.fail(404, tr(GONE))
        shown = []
        for n, item in enumerate(results):
            public = {k: v for k, v in item.items() if k not in ("file", "zipname")}
            if "file" in item:
                public["url"] = f"/api/jobs/{job.id}/docx/{n}"
            shown.append(public)
        done = sum("file" in item for item in results)
        self.reply(200, {"orientation": mode, "results": shown,
                         "zip": f"/api/jobs/{job.id}/zip" if done >= 2 else None})

    def download(self, job, n):
        with self.server.lock:
            results = job.results
        item = results[n] if n < len(results) else None
        if not item or "file" not in item:
            return self.fail(404, tr("no such result"))
        self.deliver(item["file"], DOCX, item["name"])

    def zip(self, job):
        path = job.dir / "md2docx.zip"
        with self.server.convert_lock:  # не посреди прогона и не два архива сразу
            gone = not self.current(job)  # каталог задания уже удалён — архив не собрать
            with self.server.lock:
                done = [] if gone else [item for item in job.results if "file" in item]
            if done:
                tmp = path.with_suffix(".tmp")
                with zipfile.ZipFile(tmp, "w") as z:  # docx уже сжат — храним как есть
                    for item in done:
                        z.write(item["file"], item["zipname"])
                tmp.replace(path)
        if gone:
            return self.fail(404, tr(GONE))
        if not done:
            return self.fail(404, tr("no finished files"))
        self.deliver(path, "application/zip", "md2docx.zip")


# ---------------------------------------------------------------- запуск

def port_arg(value):
    if not re.fullmatch(r"[0-9]{1,5}", value) or not 0 < int(value) < 65536:
        raise argparse.ArgumentTypeError(tr("a number from 1 to 65535 is required"))
    return int(value)


def _stop(signum, frame):
    raise SystemExit(0)


def main(argv=None, prog="md2docx_web.py"):
    parser = md2docx.Parser(prog=prog, add_help=False,
                            description=tr("Simple local md2docx interface in the browser."))
    parser._optionals.title = tr("options")
    parser.add_argument("-h", "--help", action="help", help=tr("show this help message and exit"))
    parser.add_argument("--port", type=port_arg, default=PORT, metavar="N",
                        help=tr("port on 127.0.0.1 (default {port})", port=PORT))
    parser.add_argument("--no-browser", action="store_true",
                        help=tr("do not open the browser, only print the address"))
    args = parser.parse_args(argv)
    try:
        server = Server(args.port)
    except OSError as exc:
        if exc.errno == errno.EADDRINUSE:
            print(tr("Error: port {port} is busy. If the interface is already running, it is at "
                     "http://127.0.0.1:{port}/ — otherwise check what holds the port: "
                     "lsof -nP -iTCP:{port} -sTCP:LISTEN, or start with another one: --port N",
                     port=args.port), file=sys.stderr)
        else:
            print(tr("Error: could not open port {port} ({error})", port=args.port,
                     error=exc.strerror or exc.__class__.__name__), file=sys.stderr)
        return 1
    atexit.register(server.cleanup)
    signal.signal(signal.SIGINT, signal.default_int_handler)  # даже если запущен с игнорируемым SIGINT
    for name in ("SIGTERM", "SIGHUP"):  # SIGHUP — закрыли окно Терминала
        if hasattr(signal, name):
            signal.signal(getattr(signal, name), _stop)
    url = f"http://127.0.0.1:{server.port}/"
    try:
        print(tr("md2docx v{version} — browser interface: {url}", version=md2docx.__version__, url=url),
              flush=True)
        print(tr("To stop: Ctrl+C or close this window."), flush=True)
        if not args.no_browser:
            webbrowser.open(url)
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        server.cleanup()
    try:
        print("\n" + tr("Stopped, temporary files deleted."), flush=True)
    except OSError:  # окно терминала уже закрыто
        pass
    return 0


# ---------------------------------------------------------------- страница
# Всё, что ниже, уходит в браузер: в разметке, стилях и скрипте — без комментариев.

PAGE = r"""<!doctype html>
<html lang="en" data-theme="dark">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="dark light">
<meta name="theme-color" content="#0a0a0a">
<title>md2docx — Markdown to DOCX</title>
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%2301a8a8' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z'/%3E%3Cpath d='M14 2v6h6M8 13h8M8 17h5'/%3E%3C/svg%3E">
<script nonce="__NONCE__">
try {
  const saved = localStorage.getItem("theme");
  if (saved === "light" || saved === "dark") document.documentElement.dataset.theme = saved;
} catch (e) {}
</script>
<style>
:root{--bg:#0a0a0a;--surface:#161616;--surface-2:#1f1f1f;--text:#e8e8e8;--muted:#9a9a9a;--border:#2a2a2a;--border-strong:#3c3c3c;--accent:#01a8a8;--on-accent:#0a0a0a;--accent-soft:rgba(1,168,168,.13);--ok:#4cc38a;--err:#ff6b6b;--warn:#e5b454;color-scheme:dark}
:root[data-theme="light"]{--bg:#fafaf7;--surface:#ffffff;--surface-2:#f2f2ee;--text:#1a1a1a;--muted:#6a6a6a;--border:#e0e0d8;--border-strong:#c6c6bc;--accent:#01696f;--on-accent:#ffffff;--accent-soft:rgba(1,105,111,.09);--ok:#1e7a4c;--err:#c0392b;--warn:#8a6100;color-scheme:light}
*,*::before,*::after{box-sizing:border-box}
[hidden]{display:none!important}
html{-webkit-text-size-adjust:100%;text-size-adjust:100%}
body{margin:0;background:var(--bg);color:var(--text);font:15px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;-webkit-font-smoothing:antialiased}
.wrap{max-width:760px;margin:0 auto;padding:24px 16px 32px}
.top{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:24px}
.brand{display:flex;align-items:center;gap:12px;min-width:0}
.logo{display:grid;place-items:center;flex:none;width:40px;height:40px;border-radius:10px;background:var(--accent-soft);color:var(--accent)}
h1{margin:0;font-size:20px;line-height:1.2;font-weight:700;letter-spacing:-.01em}
.sub{margin:2px 0 0;color:var(--muted);font-size:14px}
.icon{width:20px;height:20px;flex:none;fill:none;stroke:currentColor;stroke-width:2;stroke-linecap:round;stroke-linejoin:round}
.icon-btn{display:inline-grid;place-items:center;flex:none;width:40px;height:40px;padding:0;border:1px solid var(--border);border-radius:10px;background:var(--surface);color:var(--text);cursor:pointer}
.icon-btn:hover{border-color:var(--border-strong)}
.tools{display:flex;align-items:center;gap:8px;flex:none}
.lang{height:40px;max-width:8.5em;padding:0 6px;border:1px solid var(--border);border-radius:10px;background:var(--surface);color:var(--text);font:inherit;font-size:14px;cursor:pointer}
.lang:hover{border-color:var(--border-strong)}
:root[data-theme="light"] .i-sun,:root:not([data-theme="light"]) .i-moon{display:none}
:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.drop{padding:32px 16px;border:2px dashed var(--border-strong);border-radius:16px;background:var(--surface);text-align:center;transition:border-color .15s,background-color .15s}
.drop.over{border-color:var(--accent);background:var(--accent-soft)}
.drop .big{width:40px;height:40px;color:var(--accent);stroke-width:1.75}
.drop h2{margin:12px 0 4px;font-size:18px;line-height:1.3}
.drop p{max-width:460px;margin:0 auto 4px;color:var(--muted);font-size:14px}
.drop p:last-of-type{margin-bottom:18px}
.drop strong{color:var(--text)}
.btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;min-height:40px;padding:8px 16px;border:1px solid transparent;border-radius:10px;font:inherit;font-weight:600;line-height:1.2;white-space:nowrap;text-decoration:none;cursor:pointer}
.btn .icon{width:18px;height:18px}
.btn-primary{background:var(--accent);color:var(--on-accent)}
.btn-primary:hover{filter:brightness(1.08)}
.btn-ghost{background:var(--surface-2);color:var(--text);border-color:var(--border)}
.btn-ghost:hover{border-color:var(--border-strong)}
.btn[aria-disabled="true"]{opacity:.5;cursor:default}
.modes{min-width:0;margin:24px 0 0;padding:0;border:0}
.modes legend{margin-bottom:8px;padding:0;font-weight:600}
.seg{display:flex;flex-wrap:wrap;gap:6px}
.seg label{position:relative;flex:1 1 auto}
.seg input{position:absolute;top:0;left:0;width:1px;height:1px;margin:0;opacity:0}
.seg span{display:block;padding:9px 14px;border:1px solid var(--border);border-radius:10px;background:var(--surface);text-align:center;white-space:nowrap;cursor:pointer}
.seg span:hover{border-color:var(--border-strong)}
.seg input:checked+span{border-color:var(--accent);background:var(--accent-soft);box-shadow:inset 0 0 0 1px var(--accent);font-weight:600}
.seg input:focus-visible+span{outline:2px solid var(--accent);outline-offset:2px}
.hint{margin:8px 0 0;color:var(--muted);font-size:14px}
code{padding:0 4px;border:1px solid var(--border);border-radius:5px;background:var(--surface-2);font:13px/1.4 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;direction:ltr;unicode-bidi:isolate}
.status{display:flex;align-items:flex-start;gap:8px;margin:24px 0 0;overflow-wrap:anywhere}
.status:empty{display:none}
.status .icon{margin-top:1px}
.status.error{color:var(--err)}
.spin{animation:spin 1s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.actions{display:flex;flex-wrap:wrap;gap:8px;margin:12px 0 0}
.results{display:grid;gap:10px;margin:12px 0 0;padding:0;list-style:none}
.results.stale{opacity:.55}
.item{display:flex;flex-wrap:wrap;align-items:flex-start;gap:10px 16px;padding:14px 16px;border:1px solid var(--border);border-radius:12px;background:var(--surface)}
.item .main{flex:1 1 280px;min-width:0}
.name{font-weight:600;overflow-wrap:anywhere}
.line{display:flex;flex-wrap:wrap;align-items:center;gap:2px 10px;margin-top:2px}
.badge{display:inline-flex;align-items:center;gap:5px;font-size:14px;font-weight:600}
.badge .icon{width:16px;height:16px}
.badge.ok{color:var(--ok)}
.badge.err{color:var(--err)}
.badge.same{color:var(--muted)}
.meta{min-width:0;color:var(--muted);font-size:14px;overflow-wrap:anywhere}
.out{display:flex;align-items:flex-start;gap:6px;margin-top:4px;color:var(--muted);font-size:14px}
.out .icon{width:16px;height:16px;margin-top:2px}
.out span{min-width:0;overflow-wrap:anywhere}
.error-text{margin:6px 0 0;color:var(--err);font-size:14px;overflow-wrap:anywhere}
.warns{display:grid;gap:4px;margin:8px 0 0;padding:0;list-style:none}
.warns li{display:flex;align-items:flex-start;gap:8px;min-width:0;font-size:14px}
.warns .icon{width:16px;height:16px;margin-top:2px;color:var(--warn)}
.warns span{min-width:0;overflow-wrap:anywhere}
footer{display:flex;flex-wrap:wrap;justify-content:space-between;gap:4px 16px;margin-top:40px;padding-top:16px;border-top:1px solid var(--border);color:var(--muted);font-size:13px}
@media (prefers-reduced-motion:reduce){.spin{animation:none}.drop{transition:none}}
</style>
</head>
<body>
<svg width="0" height="0" style="position:absolute" aria-hidden="true" focusable="false">
<symbol id="i-file" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M8 13h8"/><path d="M8 17h5"/></symbol>
<symbol id="i-upload" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m17 8-5-5-5 5"/><path d="M12 3v12"/></symbol>
<symbol id="i-download" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5"/><path d="M12 15V3"/></symbol>
<symbol id="i-archive" viewBox="0 0 24 24"><rect x="2" y="3" width="20" height="5" rx="1"/><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8"/><path d="M10 12h4"/></symbol>
<symbol id="i-ok" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><path d="m8 12 3 3 5-6"/></symbol>
<symbol id="i-err" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><path d="M12 7v6"/><path d="M12 17h.01"/></symbol>
<symbol id="i-warn" viewBox="0 0 24 24"><path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><path d="M12 9v4"/><path d="M12 17h.01"/></symbol>
<symbol id="i-spin" viewBox="0 0 24 24"><path d="M21 12a9 9 0 1 1-6.22-8.56"/></symbol>
<symbol id="i-sun" viewBox="0 0 24 24"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M6.34 17.66l-1.41 1.41M19.07 4.93l-1.41 1.41"/></symbol>
<symbol id="i-moon" viewBox="0 0 24 24"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></symbol>
</svg>
<div class="wrap">
<header class="top">
  <div class="brand">
    <span class="logo"><svg class="icon" aria-hidden="true"><use href="#i-file"/></svg></span>
    <div>
      <h1>md2docx</h1>
      <p class="sub" data-i18n>Markdown → DOCX on this computer</p>
    </div>
  </div>
  <div class="tools">
    <select class="lang" id="lang" aria-label="Language"></select>
    <button type="button" class="icon-btn" id="theme" aria-label="Switch to light theme" title="Light theme">
      <svg class="icon i-sun" aria-hidden="true"><use href="#i-sun"/></svg>
      <svg class="icon i-moon" aria-hidden="true"><use href="#i-moon"/></svg>
    </button>
  </div>
</header>
<main>
  <section class="drop" id="drop" aria-labelledby="drop-title">
    <svg class="icon big" aria-hidden="true"><use href="#i-upload"/></svg>
    <h2 id="drop-title" data-i18n>Drop .md or .docx files here</h2>
    <p><strong dir="ltr">.md</strong> — <span data-i18n>converted to DOCX: individual files or a whole folder, so that images with relative paths are found too.</span></p>
    <p><strong dir="ltr">.docx</strong> — <span data-i18n>fonts fixed for Windows in a file that was re-saved in Word for Mac.</span></p>
    <p data-i18n>Inside a folder, .docx files are skipped — add them as individual files.</p>
    <button type="button" class="btn btn-primary" id="pick" data-i18n>Choose files</button>
    <input type="file" id="files" multiple accept=".md,.markdown,.docx,image/*" hidden>
  </section>
  <fieldset class="modes">
    <legend data-i18n>Orientation</legend>
    <div class="seg">
      <label><input type="radio" name="mode" value="auto" checked><span data-i18n>Auto</span></label>
      <label><input type="radio" name="mode" value="landscape"><span data-i18n>All landscape</span></label>
      <label><input type="radio" name="mode" value="no-auto"><span data-i18n>No auto-landscape</span></label>
    </div>
    <p class="hint" id="hint-auto" data-i18n>Wide tables move to landscape pages on their own; everything else stays portrait.</p>
    <p class="hint" id="hint-landscape" data-i18n hidden>The whole document is landscape; markers in the text still switch orientation. Same as <code>--landscape</code>.</p>
    <p class="hint" id="hint-no-auto" data-i18n hidden>Wide tables stay on portrait pages. Same as <code>--no-auto-landscape</code>.</p>
  </fieldset>
  <div class="status" id="status" role="status" aria-live="polite"></div>
  <div class="actions" id="actions" hidden>
    <a class="btn btn-primary" id="zip" href="#" download="md2docx.zip"><svg class="icon" aria-hidden="true"><use href="#i-archive"/></svg><span data-i18n>Download all (.zip)</span></a>
  </div>
  <ul class="results" id="results" aria-label="Results"></ul>
  <noscript><p>This page needs JavaScript.</p></noscript>
</main>
<footer>
  <span>md2docx v__VERSION__</span>
  <span id="where"></span>
</footer>
</div>
<script nonce="__NONCE__">
(() => {
  "use strict";
  const TOKEN = "__TOKEN__";
  const MAX_FILE = __MAX_FILE__;
  const MAX_JOB = __MAX_JOB__;
  const MB = 1048576;
  const MD_RE = /\.(md|markdown)$/i;
  const DOCX_RE = /\.docx$/i;
  const IMG_RE = /\.(png|jpe?g|gif|bmp|tiff?|webp|svg|heic|emf|wmf)$/i;
  const MODES = {"auto": "Auto", "landscape": "All landscape", "no-auto": "No auto-landscape"};
  const I18N = __I18N__;
  const LANGS = __LANGS__;
  const APP = __APP__;
  const SYSLANG = __SYSLANG__;
  const SVG = "http://www.w3.org/2000/svg";
  const $ = (id) => document.getElementById(id);
  const root = document.documentElement;
  const drop = $("drop"), input = $("files"), pick = $("pick"), statusEl = $("status");
  const list = $("results"), actions = $("actions"), zipLink = $("zip"), themeBtn = $("theme");
  const langSel = $("lang");
  const radios = Array.from(document.querySelectorAll("input[name=mode]"));
  let job = null, hasMd = false, zipUrl = null, busy = false, pending = false, depth = 0, note = null;
  let lang = "en";

  function t(key, vals) {
    const s = (I18N[lang] || {})[key] || key;
    const wrap = root.dir === "rtl" ? (v) => "⁨" + v + "⁩" : String;
    return vals ? s.replace(/\{(\w+)\}/g, (m, k) => (k in vals ? wrap(vals[k]) : m)) : s;
  }

  function preferred() {
    let saved = null;
    try { saved = localStorage.getItem("lang"); } catch (e) {}
    for (const tag of [saved, SYSLANG, ...(navigator.languages || []), navigator.language]) {
      const code = String(tag || "").toLowerCase().split(/[-_]/)[0];
      if (LANGS.some((l) => l[0] === code)) return code;
    }
    return "en";
  }

  function offline() {
    return APP ? t("the server is not responding — check that the md2docx app is running")
               : t("the server is not responding — check that the terminal window with md2docx is open");
  }

  function withNote(text) {
    return note ? t("{status}. {note}", {status: text, note: note()}) : text;
  }

  function icon(name, cls) {
    const s = document.createElementNS(SVG, "svg");
    s.setAttribute("class", "icon" + (cls ? " " + cls : ""));
    s.setAttribute("aria-hidden", "true");
    const u = document.createElementNS(SVG, "use");
    u.setAttribute("href", "#" + name);
    s.append(u);
    return s;
  }

  function el(tag, cls, text) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }

  function setStatus(text, kind) {
    statusEl.className = "status" + (kind ? " " + kind : "");
    statusEl.replaceChildren();
    if (!text) return;
    if (kind === "busy") statusEl.append(icon("i-spin", "spin"));
    if (kind === "error") statusEl.append(icon("i-err"));
    statusEl.append(el("span", null, text));
  }

  function setBusy(on) {
    busy = on;
    pick.setAttribute("aria-disabled", on ? "true" : "false");
    list.classList.toggle("stale", on);
  }

  function mode() {
    return radios.find((r) => r.checked).value;
  }

  function showHint() {
    for (const key of Object.keys(MODES)) $("hint-" + key).hidden = key !== mode();
  }

  async function request(method, url, body, type) {
    const headers = {"X-Token": TOKEN, "X-Lang": lang};
    if (type) headers["Content-Type"] = type;
    let r;
    try {
      r = await fetch(url, {method, headers, body});
    } catch (e) {
      throw new Error(offline());
    }
    if (r.ok) return r;
    let data = {};
    try { data = await r.json(); } catch (e) {}
    const err = new Error(data.error || t("server responded with {status}", {status: r.status}));
    err.status = r.status;
    throw err;
  }

  async function api(method, url, body, type) {
    const r = await request(method, url, body, type);
    try { return await r.json(); } catch (e) { return {}; }
  }

  function stale() {
    list.classList.add("stale");
    actions.hidden = true;
  }

  async function save(e, url, name) {
    e.preventDefault();
    const a = e.currentTarget;
    if (busy || a.getAttribute("aria-disabled") === "true") return;
    a.setAttribute("aria-disabled", "true");
    try {
      if (APP) {
        const r = await request("POST", url);
        let data;
        try { data = await r.json(); } catch (err) { throw new Error(offline()); }
        if (data.error) throw new Error(data.error);
        return;
      }
      const r = await request("GET", url);
      let blob;
      try { blob = await r.blob(); } catch (err) { throw new Error(offline()); }
      const link = el("a");
      link.href = URL.createObjectURL(blob);
      link.download = name;
      document.body.append(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(link.href), 60000);
    } catch (err) {
      setStatus(t("Could not download: {error}", {error: err.message}), "error");
      if (err.status === 404) stale();
    } finally {
      a.removeAttribute("aria-disabled");
    }
  }

  function hiddenPath(path) {
    return path.split("/").some((part) => part.startsWith("."));
  }

  function keep(path) {
    if (hiddenPath(path)) return false;
    if (DOCX_RE.test(path)) return !path.includes("/");
    return MD_RE.test(path) || IMG_RE.test(path);
  }

  function readBatch(reader) {
    return new Promise((resolve, reject) => reader.readEntries(resolve, reject));
  }

  function entryFile(entry) {
    return new Promise((resolve, reject) => entry.file(resolve, reject));
  }

  async function walk(entry, out) {
    if (entry.name.startsWith(".")) return;
    if (entry.isFile) {
      const path = entry.fullPath.replace(/^\/+/, "");
      if (keep(path)) out.push({path, file: await entryFile(entry)});
    } else if (entry.isDirectory) {
      const reader = entry.createReader();
      for (let batch = await readBatch(reader); batch.length; batch = await readBatch(reader)) {
        for (const child of batch) await walk(child, out);
      }
    }
  }

  async function fromDrop(dt) {
    const entries = [], out = [];
    for (const item of Array.from(dt.items || [])) {
      if (item.kind !== "file") continue;
      const entry = item.webkitGetAsEntry ? item.webkitGetAsEntry() : null;
      if (entry) entries.push(entry);
      else {
        const f = item.getAsFile();
        if (f) out.push({path: f.name, file: f});
      }
    }
    if (!dt.items) for (const f of Array.from(dt.files || [])) out.push({path: f.name, file: f});
    for (const entry of entries) await walk(entry, out);
    return out;
  }

  function unique(files) {
    const seen = new Set();
    return files.map((f) => {
      let path = f.path, n = 1;
      while (seen.has(path.toLowerCase())) {
        n += 1;
        path = f.path.replace(/(\.[^./]*)?$/, " (" + n + ")$1");
      }
      seen.add(path.toLowerCase());
      return {path, file: f.file};
    });
  }

  async function start(picked) {
    let files = unique(picked.filter((f) => keep(f.path)));
    files.sort((a, b) => a.path.localeCompare(b.path, lang, {numeric: true}));
    const big = files.filter((f) => f.file.size > MAX_FILE);
    files = files.filter((f) => f.file.size <= MAX_FILE);
    note = big.length ? () => t("Skipped files larger than {size} MB: {count} ({files})",
      {size: MAX_FILE / MB, count: big.length, files: big.map((f) => f.path).join(", ")}) : null;
    if (!files.some((f) => MD_RE.test(f.path) || DOCX_RE.test(f.path))) {
      setStatus(withNote(t("No .md or .docx files among the selected ones (.docx files inside folders are skipped)")), "error");
      return;
    }
    const total = files.reduce((sum, f) => sum + f.file.size, 0);
    if (total > MAX_JOB) {
      setStatus(t("The files together exceed {size} MB — choose fewer files or a smaller folder", {size: MAX_JOB / MB}), "error");
      return;
    }
    job = null;
    hasMd = files.some((f) => MD_RE.test(f.path));
    list.replaceChildren();
    actions.hidden = true;
    setBusy(true);
    try {
      const created = await api("POST", "/api/jobs");
      for (let i = 0; i < files.length; i++) {
        setStatus(t("Uploading {n} of {total}: {file}", {n: i + 1, total: files.length, file: files[i].path}), "busy");
        await api("PUT", "/api/jobs/" + created.job + "/files?path=" + encodeURIComponent(files[i].path), files[i].file);
      }
      job = created.job;
    } catch (e) {
      setStatus(t("Could not upload: {error}", {error: e.message}), "error");
      return;
    } finally {
      setBusy(false);
    }
    await convert();
  }

  async function convert() {
    if (!job) return;
    if (busy) {
      pending = true;
      return;
    }
    pending = false;
    const m = mode();
    let gone = false;
    setBusy(true);
    setStatus(hasMd ? t("Converting ({mode})…", {mode: t(MODES[m])}) : t("Fixing fonts…"), "busy");
    try {
      const data = await api("POST", "/api/jobs/" + job + "/convert", JSON.stringify({orientation: m}), "application/json");
      render(data, m);
    } catch (e) {
      setStatus(hasMd ? t("Could not convert: {error}", {error: e.message})
                      : t("Could not fix fonts: {error}", {error: e.message}), "error");
      gone = e.status === 404;
    } finally {
      setBusy(false);
    }
    if (gone) stale();
    if (pending) {
      pending = false;
      await convert();
    }
  }

  function item(r) {
    const li = el("li", "item");
    const main = el("div", "main");
    const source = el("div", "name", r.source);
    source.dir = "auto";
    main.append(source);
    const line = el("div", "line");
    const fonts = r.kind === "docx";
    const label = !r.ok ? t("error") : !fonts ? t("done") : r.fixed ? t("fonts fixed") : t("nothing to fix — the hints are in place");
    const badge = el("span", "badge " + (!r.ok ? "err" : fonts && !r.fixed ? "same" : "ok"));
    badge.append(icon(r.ok ? "i-ok" : "i-err"), el("span", null, label));
    line.append(badge);
    if (r.ok && r.summary) line.append(el("span", "meta", r.summary));
    main.append(line);
    if (!r.ok) main.append(el("p", "error-text", r.error));
    if (r.url) {
      const out = el("div", "out");
      const name = el("span", null, r.name);
      name.dir = "auto";
      out.append(icon("i-file"), name);
      main.append(out);
    }
    if (r.warnings.length) {
      const warns = el("ul", "warns");
      warns.setAttribute("aria-label", t("Warnings"));
      for (const w of r.warnings) {
        const wl = el("li");
        wl.append(icon("i-warn"), el("span", null, w));
        warns.append(wl);
      }
      main.append(warns);
    }
    li.append(main);
    if (r.url) {
      const a = el("a", "btn btn-ghost");
      a.href = r.url + "?token=" + encodeURIComponent(TOKEN);
      a.download = r.name;
      a.setAttribute("aria-label", t("Download {file}", {file: r.name}));
      a.append(icon("i-download"), el("span", null, t("Download .docx")));
      a.addEventListener("click", (e) => save(e, r.url, r.name));
      li.append(a);
    }
    return li;
  }

  function render(data, m) {
    list.replaceChildren(...data.results.map(item));
    const ok = data.results.filter((r) => r.ok).length;
    const failed = data.results.length - ok;
    zipUrl = data.zip;
    actions.hidden = !zipUrl;
    if (zipUrl) zipLink.href = zipUrl + "?token=" + encodeURIComponent(TOKEN);
    const vals = {mode: t(MODES[m]), ok, total: data.results.length, failed};
    setStatus(withNote(hasMd ? (failed ? t("{mode}: done {ok} of {total}, failed {failed}", vals) : t("{mode}: done {ok} of {total}", vals))
                             : (failed ? t("Done {ok} of {total}, failed {failed}", vals) : t("Done {ok} of {total}", vals))));
  }

  function isFileDrag(e) {
    return !!e.dataTransfer && Array.from(e.dataTransfer.types || []).includes("Files");
  }

  window.addEventListener("dragenter", (e) => {
    if (!isFileDrag(e)) return;
    e.preventDefault();
    depth += 1;
    drop.classList.add("over");
  });
  window.addEventListener("dragover", (e) => {
    if (!isFileDrag(e)) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = busy ? "none" : "copy";
  });
  window.addEventListener("dragleave", (e) => {
    if (!isFileDrag(e)) return;
    depth = Math.max(0, depth - 1);
    if (!depth) drop.classList.remove("over");
  });
  window.addEventListener("drop", (e) => {
    if (!isFileDrag(e)) return;
    e.preventDefault();
    depth = 0;
    drop.classList.remove("over");
    if (busy) return;
    setBusy(true);
    setStatus(t("Reading the dropped items…"), "busy");
    fromDrop(e.dataTransfer).then((files) => {
      setBusy(false);
      return start(files);
    }, (err) => {
      setBusy(false);
      setStatus(t("Could not read the dropped items: {error}", {error: err && err.message ? err.message : err}), "error");
    });
  });

  pick.addEventListener("click", () => {
    if (!busy) input.click();
  });
  input.addEventListener("change", () => {
    const files = Array.from(input.files).map((f) => ({path: f.webkitRelativePath || f.name, file: f}));
    input.value = "";
    if (files.length) start(files);
  });
  zipLink.addEventListener("click", (e) => save(e, zipUrl, "md2docx.zip"));
  for (const r of radios) {
    r.addEventListener("change", () => {
      showHint();
      if (hasMd) convert();
    });
  }

  function paintTheme() {
    const light = root.dataset.theme === "light";
    themeBtn.setAttribute("aria-label", light ? t("Switch to dark theme") : t("Switch to light theme"));
    themeBtn.title = light ? t("Dark theme") : t("Light theme");
    document.querySelector("meta[name=theme-color]").content = light ? "#fafaf7" : "#0a0a0a";
  }
  themeBtn.addEventListener("click", () => {
    root.dataset.theme = root.dataset.theme === "light" ? "dark" : "light";
    try { localStorage.setItem("theme", root.dataset.theme); } catch (e) {}
    paintTheme();
  });

  const texts = Array.from(document.querySelectorAll("[data-i18n]"), (e) => {
    const kids = [];
    let key = "";
    for (const n of e.childNodes) {
      if (n.nodeType === Node.TEXT_NODE) key += n.data;
      else key += "{" + (kids.push(n) - 1) + "}";
    }
    return {e, key, kids};
  });

  function translate() {
    root.lang = lang;
    root.dir = lang === "ar" ? "rtl" : "ltr";
    document.title = t("md2docx — Markdown to DOCX");
    for (const {e, key, kids} of texts) {
      e.replaceChildren(...t(key).split(/\{(\d+)\}/).map((part, i) => (i % 2 ? kids[part] : part)));
    }
    langSel.setAttribute("aria-label", t("Language"));
    list.setAttribute("aria-label", t("Results"));
    $("where").textContent = APP ? t("Runs only on this computer. To stop it, quit the md2docx app.")
                                 : t("Runs only on this computer. To stop it, press Ctrl+C in the terminal window.");
    paintTheme();
  }

  for (const [code, name] of LANGS) {
    const o = new Option(name, code);
    o.lang = code;
    langSel.append(o);
  }
  lang = langSel.value = preferred();
  langSel.addEventListener("change", () => {
    lang = langSel.value;
    try { localStorage.setItem("lang", lang); } catch (e) {}
    translate();
    if (job) convert();
  });
  translate();
  showHint();
})();
</script>
</body>
</html>
"""


def i18n_key(inner):
    """Текст элемента с data-i18n → ключ перевода, как его собирает скрипт страницы:
    вложенные элементы заменяются на {0}, {1}…"""
    parts = re.split(r"<(\w+)[^>]*>.*?</\1>", inner, flags=re.S)
    return "".join(f"{{{i // 2}}}" if i % 2 else part for i, part in enumerate(parts))


# строки страницы, которым нужен перевод: тексты элементов с data-i18n и литералы t("…") в скрипте
PAGE_KEYS = frozenset(
    [i18n_key(m[2]) for m in re.finditer(r"<(\w+)[^>]*\bdata-i18n\b[^>]*>(.*?)</\1>", PAGE, re.S)]
    + [json.loads(f'"{s}"') for s in re.findall(r'\bt\("((?:[^"\\]|\\.)*)"', PAGE)])


def script_json(obj):
    """JSON для вставки в <script>: «<» экранирован, чтобы «</script>» в строке не закрыл тег."""
    return json.dumps(obj, ensure_ascii=False).replace("<", "\\u003c")


@functools.cache
def page_catalogs():
    """Переводы строк страницы на все языки, кроме английского: {язык: {строка: перевод}}."""
    return script_json({lang: {k: v for k, v in md2docx.catalog(lang).items() if k in PAGE_KEYS}
                        for lang in md2docx.LANGUAGES if lang != "en"})


if __name__ == "__main__":
    sys.exit(main())
