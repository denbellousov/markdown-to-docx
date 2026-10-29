#!/usr/bin/env python3
"""Независимый сверщик md ↔ docx для md2docx (SPEC.md §6).

    python3 tests/verify.py IN.md OUT.docx [--expect-orient P,L,P]
    python3 tests/verify.py --invariants-only OUT.docx

Написан по SPEC.md как по чёрному ящику: md2docx.py не импортируется и не
читается. Markdown разбирается своим кодом поверх markdown-it-py с той же
конфигурацией, что в SPEC §2.

Что сверяется (полный режим):
  * слова — мультимножество и порядок (проекция md без маркеров, комментариев и
    front matter; URL ссылок не входят; картинка — либо рисунок, либо заглушка
    «[изображение: alt]», текст заглушки из сверки слов исключён, а её alt
    сверяется с md отдельно);
  * последовательность заголовков (уровень + текст), таблицы (число, размеры,
    тексты ячеек, выравнивание колонок, таблица без шапки), остатки разметки,
    нумерация списков, адреса гиперссылок по порядку;
  * план разделов по md (§4): маркеры, разрывы, автоальбомные группы с правилом
    подтягивания и склейкой → ориентация и границы разделов, разрывы страниц;
  * инварианты docx (работают и в режиме --invariants-only): §3.1 страница,
    §3.2 docDefaults / settings / fontTable / свойства документа / прогоны Arial
    для символов без глифа в HN, §3.3 стили заголовков, §3.4 таблицы (ширины по
    жёсткому и мягкому min, keepNext, отбивка после заголовка, хвостовой абзац),
    §3.5 уровни нумерации, §3.6 схемы гиперссылок и размер рисунков, §3.8 глубина
    отступов, §4.3 широкие таблицы, §4.4 разделы и разрывы, §5 порядок элементов;
  * целостность пакета (§6(в), то, на что Word отвечает «файл повреждён»): типы
    содержимого всех частей, внутренние связи, префиксы mc:Ignorable, id рисунков.

Код выхода: 0 — расхождений нет; 1 — есть расхождения (или файл не читается);
2 — ошибка аргументов.
"""
from __future__ import annotations

import argparse
import io
import re
import sys
import traceback
import unicodedata
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree
from markdown_it import MarkdownIt
from markdown_it.common.utils import unescapeAll
from markdown_it.rules_block import reference as _md_reference

# --------------------------------------------------------------------------
# Пространства имён и мелкие помощники
# --------------------------------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
W = "{%s}" % W_NS
NS = {"w": W_NS, "r": R_NS}


def q(tag: str) -> str:
    return W + tag


def wa(el, name: str, default=None):
    """Атрибут w:<name> элемента; default, если элемента или атрибута нет."""
    if el is None:
        return default
    return el.get(W + name, default)


def to_int(v):
    if v is None:
        return None
    try:
        return int(v)
    except ValueError:
        try:
            return int(round(float(v)))
        except ValueError:
            return None


def is_on(el) -> bool:
    """Переключатель OOXML (w:b, w:i, w:keepNext …): есть и не выключен."""
    if el is None:
        return False
    v = el.get(W + "val")
    return v is None or v.lower() not in ("0", "false", "off", "none")


def local(el) -> str | None:
    """Локальное имя элемента из пространства w:, иначе None."""
    tag = el.tag
    if not isinstance(tag, str) or not tag.startswith(W):
        return None
    return tag[len(W):]


def norm_ws(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def norm_lines(s: str) -> tuple:
    """Текст → кортеж строк (разрез по \\n), пробелы схлопнуты, края без пустых."""
    lines = [norm_ws(x) for x in s.split("\n")]
    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()
    return tuple(lines)


def short(s: str, n: int = 60) -> str:
    s = s.replace("\n", " ⏎ ")
    return s if len(s) <= n else s[: n - 1] + "…"


# --------------------------------------------------------------------------
# Константы спецификации
# --------------------------------------------------------------------------

A4 = {"P": (11906, 16838), "L": (16838, 11906)}
MARGIN = 1134
HDR_FTR = 720
TEXT_WIDTH = {"P": 9638, "L": 14570}
TEXT_HEIGHT = {"P": 14570, "L": 9638}
ROW_LINE = 262  # шаг строки ячейки, SPEC §3.4 «Строки»
RU_ORIENT = {"P": "книжный", "L": "альбомный"}
ORIENT_ALIASES = {"P": "P", "L": "L", "К": "P", "А": "L", "K": "P", "A": "L",
                  "PORTRAIT": "P", "LANDSCAPE": "L"}
HN = "Helvetica Neue"
HN_PANOSE = "02000503000000020004"
HEADING_SPEC = {1: (36, 360), 2: (32, 300), 3: (28, 260),
                4: (24, 240), 5: (22, 220), 6: (22, 200)}
BORDER_SIDES = ("top", "left", "bottom", "right", "insideH", "insideV")
SIDE_ALIASES = {"left": ("left", "start"), "right": ("right", "end")}
CELL_MAR = {"top": 40, "bottom": 40, "left": 108, "right": 108}
MAX_TBL_IND = 9 * 360  # §3.8: отступ не глубже 8 уровней вложенности (+ отступ пункта списка)
MAX_PARA_IND = 10 * 360
TBL_LOOK = {"val": "04A0", "firstRow": "1", "lastRow": "0", "firstColumn": "1",
            "lastColumn": "0", "noHBand": "0", "noVBand": "1"}
BULLETS = ("•", "◦", "▪")
JUSTIFY = {"both", "distribute", "lowKashida", "mediumKashida", "highKashida",
           "thaiDistribute"}
THEME_FONT_ATTRS = ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme")
CODE_FONTS = {"Menlo", "Consolas", "Courier New"}
PLACEHOLDER_RE = re.compile(r"\[изображение(?::[^\]]*)?\]")

# --------------------------------------------------------------------------
# §5 — порядок дочерних элементов (ECMA-376). Кортеж = равноправные имена.
# --------------------------------------------------------------------------

ORDER = {
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
            "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
            "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
            "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
            "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
            "suppressOverlap", "jc", "textDirection", "textAlignment",
            "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
            "pPrChange"],
    "rPr": ["ins", "del", "moveFrom", "moveTo",
            "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
            "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
            "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
            "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
            "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
            "specVanish", "oMath", "rPrChange"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
              "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
              "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
              "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange"],
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd",
             "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark",
             "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"],
    "sectPr": [("headerReference", "footerReference"), "footnotePr", "endnotePr",
               "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType",
               "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg",
               "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings",
               "sectPrChange"],
    # Не перечислены в SPEC §5, но правило то же (схема ECMA-376, Word строг):
    "tblBorders": ["top", ("left", "start"), "bottom", ("right", "end"),
                   "insideH", "insideV"],
    "tcBorders": ["top", ("left", "start"), "bottom", ("right", "end"),
                  "insideH", "insideV", "tl2br", "tr2bl"],
    "pBdr": ["top", ("left", "start"), "bottom", ("right", "end"), "between", "bar"],
    "tblCellMar": ["top", ("left", "start"), "bottom", ("right", "end")],
    "tcMar": ["top", ("left", "start"), "bottom", ("right", "end")],
    "numPr": ["ilvl", "numId", "numberingChange", "ins"],
    "style": ["name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden",
              "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked",
              "personal", "personalCompose", "personalReply", "rsid", "pPr", "rPr",
              "tblPr", "trPr", "tcPr", "tblStylePr"],
    "docDefaults": ["rPrDefault", "pPrDefault"],
    "font": ["altName", "panose1", "charset", "family", "notTrueType", "pitch", "sig",
             "embedRegular", "embedBold", "embedItalic", "embedBoldItalic"],
}
NUMBERING_ORDER = {
    "numbering": ["numPicBullet", "abstractNum", "num", "numIdMacAtCleanup"],
    "abstractNum": ["nsid", "multiLevelType", "tmpl", "name", "styleLink",
                    "numStyleLink", "lvl"],
    "lvl": ["start", "numFmt", "lvlRestart", "pStyle", "isLgl", "suff", "lvlText",
            "lvlPicBulletId", "legacy", "lvlJc", "pPr", "rPr"],
    "num": ["abstractNumId", "lvlOverride"],
    "lvlOverride": ["startOverride", "lvl"],
}
REPEATABLE = {"headerReference", "footerReference", "tblStylePr", "lvl",
              "lvlOverride", "numPicBullet", "abstractNum", "num"}


def _rank_table(order):
    rank = {}
    for i, item in enumerate(order):
        for name in (item if isinstance(item, tuple) else (item,)):
            rank[name] = i
    return rank


ORDER_RANK = {k: _rank_table(v) for k, v in ORDER.items()}
NUMBERING_RANK = {k: _rank_table(v) for k, v in NUMBERING_ORDER.items()}

# --------------------------------------------------------------------------
# Отчёт
# --------------------------------------------------------------------------

CATEGORY_TITLES = {
    "words": "слова", "headings": "заголовки", "tables": "таблицы (содержимое)",
    "markup": "остатки разметки", "images": "картинки", "lists": "списки",
    "page": "§3.1 страница", "defaults": "§3.2 docDefaults/settings",
    "fonts": "§3.2 fontTable", "core": "§3.2 свойства документа",
    "heading-styles": "§3.3 заголовки", "table-format": "§3.4 таблицы",
    "numbering": "§3.5/§5 нумерация", "auto-landscape": "§4.3 автоальбомная",
    "sections": "§4.4 разделы", "orient": "ориентация разделов",
    "order": "§5 порядок элементов", "structure": "структура OOXML",
    "package": "пакет OOXML (типы, связи, mc:Ignorable)",
    "links": "§3.6 ссылки", "glyphs": "§3.2 символы без глифа в HN",
    "indent": "§3.8 отступы", "plan": "§4 план разделов по md", "input": "§2 вход md",
    "internal": "сбой сверщика",
}


@dataclass
class Issue:
    code: str
    msg: str

    def __str__(self):
        return f"[{CATEGORY_TITLES.get(self.code, self.code)}] {self.msg}"


@dataclass
class Report:
    md_path: str | None = None
    docx_path: str | None = None
    errors: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    info: list = field(default_factory=list)

    def err(self, code, msg):
        self.errors.append(Issue(code, msg))

    def warn(self, code, msg):
        self.warnings.append(Issue(code, msg))

    @property
    def ok(self) -> bool:
        return not self.errors

    def codes(self) -> set:
        return {e.code for e in self.errors}

    def has(self, code: str, needle: str = "") -> bool:
        return any(e.code == code and needle in e.msg for e in self.errors)

    def render(self) -> str:
        out = []
        head = f"Сверка: {self.md_path} ↔ {self.docx_path}" if self.md_path else \
            f"Инварианты: {self.docx_path} (без сверки с md)"
        out.append(head)
        out.extend("  " + s for s in self.info)
        if self.errors:
            out.append(f"Расхождения: {len(self.errors)}")
            for i, e in enumerate(self.errors, 1):
                out.append(f"  {i}. {e}")
        if self.warnings:
            out.append(f"Предупреждения: {len(self.warnings)}")
            for e in self.warnings:
                out.append(f"  - {e}")
        n = len(self.errors)
        out.append("ИТОГ: OK — расхождений нет" if not n else
                   f"ИТОГ: FAIL — расхождений {n}")
        return "\n".join(out)


# --------------------------------------------------------------------------
# Замер ширины текста (SPEC §3.4 п. 2) — для инварианта «слова не рвутся»
# --------------------------------------------------------------------------

_HN_FILE = "/System/Library/Fonts/HelveticaNeue.ttc"
_MENLO_FILE = "/System/Library/Fonts/Menlo.ttc"
_TTC_INDEX = {(False, False): 0, (True, False): 1, (False, True): 2, (True, True): 3}
_ARIAL = {
    (False, False): ["/System/Library/Fonts/Supplemental/Arial.ttf", "C:/Windows/Fonts/arial.ttf"],
    (True, False): ["/System/Library/Fonts/Supplemental/Arial Bold.ttf", "C:/Windows/Fonts/arialbd.ttf"],
    (False, True): ["/System/Library/Fonts/Supplemental/Arial Italic.ttf", "C:/Windows/Fonts/ariali.ttf"],
    (True, True): ["/System/Library/Fonts/Supplemental/Arial Bold Italic.ttf", "C:/Windows/Fonts/arialbi.ttf"],
}
_COURIER = {
    (False, False): ["/System/Library/Fonts/Supplemental/Courier New.ttf", "C:/Windows/Fonts/cour.ttf"],
    (True, False): ["/System/Library/Fonts/Supplemental/Courier New Bold.ttf", "C:/Windows/Fonts/courbd.ttf"],
    (False, True): ["/System/Library/Fonts/Supplemental/Courier New Italic.ttf", "C:/Windows/Fonts/couri.ttf"],
    (True, True): ["/System/Library/Fonts/Supplemental/Courier New Bold Italic.ttf", "C:/Windows/Fonts/courbi.ttf"],
}


class Measurer:
    """Ширина строки в twips. Pillow по реальному шрифту, крупный кегль
    (без хинтинга) с пересчётом; запасной путь — Arial, затем эвристика."""

    REF = 1000

    def __init__(self):
        self._cache = {}
        try:
            from PIL import ImageFont  # noqa: F401
            self._pil = True
        except Exception:  # pragma: no cover
            self._pil = False
        self.source = "Pillow/" + ("Helvetica Neue" if Path(_HN_FILE).exists() else "Arial") \
            if self._pil else "эвристика"

    def _font(self, code: bool, bold: bool, italic: bool):
        key = (code, bold, italic)
        if key in self._cache:
            return self._cache[key]
        font = None
        if self._pil:
            from PIL import ImageFont
            main = _MENLO_FILE if code else _HN_FILE
            if Path(main).exists():
                try:
                    font = ImageFont.truetype(main, self.REF, index=_TTC_INDEX[(bold, italic)])
                except OSError:
                    font = None
            if font is None:
                for p in (_COURIER if code else _ARIAL)[(bold, italic)]:
                    if Path(p).exists():
                        try:
                            font = ImageFont.truetype(p, self.REF)
                            break
                        except OSError:
                            pass
        self._cache[key] = font
        return font

    def _arial(self, bold: bool, italic: bool):
        """Arial — замена HN в Word на Windows (SPEC §3.4 п. 2: ширина — бо́льшая из двух)."""
        key = ("arial", bold, italic)
        if key not in self._cache:
            self._cache[key] = None
            if self._pil:
                from PIL import ImageFont
                for p in _ARIAL[(bold, italic)]:
                    if Path(p).exists():
                        try:
                            self._cache[key] = ImageFont.truetype(p, self.REF)
                            break
                        except OSError:
                            pass
        return self._cache[key]

    def twips(self, text: str, size_pt: float, code=False, bold=False, italic=False) -> float:
        if not text:
            return 0.0
        f = self._font(code, bold, italic)
        if f is None:
            em = 0.60 if (bold or code) else 0.55
            return len(text) * em * size_pt * 20
        width = f.getlength(text)
        a = None if code else self._arial(bold, italic)
        if a is not None:
            width = max(width, a.getlength(text))
        return width * size_pt / self.REF * 20


_MEASURER = None


def measurer() -> Measurer:
    global _MEASURER
    if _MEASURER is None:
        _MEASURER = Measurer()
    return _MEASURER


# --------------------------------------------------------------------------
# Модель DOCX
# --------------------------------------------------------------------------


@dataclass
class Section:
    index: int
    sectPr: object
    elements: list
    carrier: object  # w:p с sectPr (None у последнего раздела)

    def _pg(self):
        return self.sectPr.find(q("pgSz")) if self.sectPr is not None else None

    @property
    def orient(self) -> str:
        pg = self._pg()
        if pg is None:
            return "?"
        if wa(pg, "orient") == "landscape":
            return "L"
        w_, h_ = to_int(wa(pg, "w")), to_int(wa(pg, "h"))
        if wa(pg, "orient") is None and w_ and h_ and w_ > h_:
            return "L"  # без атрибута, но шире высоты — так рисует Word
        return "P"

    @property
    def text_width(self) -> int | None:
        pg, mar = self._pg(), self.sectPr.find(q("pgMar")) if self.sectPr is not None else None
        if pg is None or mar is None:
            return None
        w_ = to_int(wa(pg, "w"))
        left, right = to_int(wa(mar, "left")) or 0, to_int(wa(mar, "right")) or 0
        gutter = to_int(wa(mar, "gutter")) or 0
        return None if w_ is None else w_ - left - right - gutter


@dataclass
class DTable:
    el: object
    index: int
    section: int
    grid: list
    rows: list  # список списков w:tc


class Docx:
    def __init__(self, path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path) as z:
            self.parts = {n: z.read(n) for n in z.namelist()}
        self.document = self._xml("word/document.xml")
        if self.document is None:
            raise ValueError("в пакете нет word/document.xml")
        self.styles = self._xml("word/styles.xml")
        self.numbering = self._xml("word/numbering.xml")
        self.settings = self._xml("word/settings.xml")
        self.font_table = self._xml("word/fontTable.xml")
        self.core = self._xml("docProps/core.xml")
        self.app = self._xml("docProps/app.xml")
        self.rels = self._xml("word/_rels/document.xml.rels")
        self.body = self.document.find(q("body"))
        if self.body is None:
            raise ValueError("в document.xml нет w:body")
        self.styles_by_id = {}
        self.default_pstyle = None
        self.ppr_default = self.rpr_default = None
        if self.styles is not None:
            for s in self.styles.findall(q("style")):
                self.styles_by_id[wa(s, "styleId")] = s
                if wa(s, "type") == "paragraph" and wa(s, "default") in ("1", "true"):
                    self.default_pstyle = wa(s, "styleId")
            if self.default_pstyle is None and "Normal" in self.styles_by_id:
                self.default_pstyle = "Normal"
            dd = self.styles.find(q("docDefaults"))
            if dd is not None:
                self.ppr_default = dd.find("w:pPrDefault/w:pPr", NS)
                self.rpr_default = dd.find("w:rPrDefault/w:rPr", NS)
        self.rel_by_id = {}
        if self.rels is not None:
            for r in self.rels:
                if isinstance(r.tag, str):
                    self.rel_by_id[r.get("Id")] = r
        self.sections, self.section_of = self._build_sections()
        self.tables = self._build_tables()

    def _xml(self, name):
        data = self.parts.get(name)
        if data is None:
            return None
        return etree.fromstring(data)

    # --- разделы ---------------------------------------------------------
    def _build_sections(self):
        sections, section_of, cur = [], {}, []
        final = None
        for ch in self.body:
            if not isinstance(ch.tag, str):
                continue
            if ch.tag == q("sectPr"):
                final = ch
                continue
            cur.append(ch)
            section_of[ch] = len(sections)
            if ch.tag == q("p"):
                sp = ch.find("w:pPr/w:sectPr", NS)
                if sp is not None:
                    sections.append(Section(len(sections), sp, cur, ch))
                    cur = []
        sections.append(Section(len(sections), final, cur, None))
        return sections, section_of

    def top_of(self, el):
        """Элемент верхнего уровня body, содержащий el."""
        while el is not None and el.getparent() is not self.body:
            el = el.getparent()
        return el

    def section_index(self, el) -> int:
        return self.section_of.get(self.top_of(el), len(self.sections) - 1)

    # --- таблицы ---------------------------------------------------------
    def _build_tables(self):
        out = []
        for i, tbl in enumerate(self.body.iter(q("tbl"))):
            grid = [to_int(wa(g, "w")) or 0 for g in tbl.findall("w:tblGrid/w:gridCol", NS)]
            rows = [tr.findall(q("tc")) for tr in tbl.findall(q("tr"))]
            out.append(DTable(tbl, i + 1, self.section_index(tbl), grid, rows))
        return out

    # --- стили -----------------------------------------------------------
    def style_chain(self, sid):
        out, seen = [], set()
        while sid and sid not in seen:
            seen.add(sid)
            st = self.styles_by_id.get(sid)
            if st is None:
                break
            out.append(st)
            sid = wa(st.find(q("basedOn")), "val")
        return out

    def pstyle_id(self, p):
        return wa(p.find("w:pPr/w:pStyle", NS), "val") or self.default_pstyle

    def style_heading_level(self, sid):
        m = re.fullmatch(r"Heading([1-9])", sid or "")
        if m:
            return int(m.group(1))
        st = self.styles_by_id.get(sid)
        if st is not None:
            name = (wa(st.find(q("name")), "val") or "").strip().lower()
            m = re.fullmatch(r"heading ([1-9])", name)
            if m:
                return int(m.group(1))
        return None

    def heading_level(self, p):
        ps = p.find("w:pPr/w:pStyle", NS)
        return None if ps is None else self.style_heading_level(wa(ps, "val"))

    def find_style_by_name(self, name):
        for s in self.styles_by_id.values():
            if (wa(s.find(q("name")), "val") or "").strip().lower() == name.lower():
                return s
        return None

    def eff_ppr(self, p, child, attr=None):
        """Значение свойства абзаца: pPr → цепочка стиля → docDefaults."""
        sources = [p.find(q("pPr"))]
        sources += [st.find(q("pPr")) for st in self.style_chain(self.pstyle_id(p))]
        sources.append(self.ppr_default)
        for src in sources:
            if src is None:
                continue
            el = src.find(q(child))
            if el is None:
                continue
            if attr is None:
                return el
            v = wa(el, attr)
            if v is not None:
                return v
        return None

    def eff_rpr(self, r, p, child):
        """Элемент свойства прогона: rPr → rStyle → стиль абзаца → docDefaults."""
        rpr = r.find(q("rPr"))
        sources = [rpr]
        rs = wa(rpr.find(q("rStyle")), "val") if rpr is not None else None
        sources += [st.find(q("rPr")) for st in self.style_chain(rs)]
        if p is not None:
            sources += [st.find(q("rPr")) for st in self.style_chain(self.pstyle_id(p))]
        sources.append(self.rpr_default)
        for src in sources:
            if src is not None:
                el = src.find(q(child))
                if el is not None:
                    return el
        return None

    def eff_font(self, r, p, attr="ascii"):
        rpr = r.find(q("rPr"))
        rs = wa(rpr.find(q("rStyle")), "val") if rpr is not None else None
        sources = [rpr] + [st.find(q("rPr")) for st in self.style_chain(rs)]
        if p is not None:
            sources += [st.find(q("rPr")) for st in self.style_chain(self.pstyle_id(p))]
        sources.append(self.rpr_default)
        return font_from_sources(sources, attr)


def font_from_sources(sources, attr="ascii"):
    theme_attr = {"ascii": "asciiTheme", "hAnsi": "hAnsiTheme",
                  "eastAsia": "eastAsiaTheme", "cs": "cstheme"}[attr]
    for src in sources:
        if src is None:
            continue
        rf = src.find(q("rFonts"))
        if rf is None:
            continue
        if wa(rf, theme_attr) is not None:
            return "theme:" + wa(rf, theme_attr)
        if wa(rf, attr) is not None:
            return wa(rf, attr)
    return None


def para_text(p) -> str:
    """Видимый текст абзаца: w:t, w:tab → \\t, w:br/w:cr → \\n."""
    out = []
    for el in p.iter():
        tag = el.tag
        if tag == q("t"):
            out.append(el.text or "")
        elif tag == q("tab"):
            par = el.getparent()
            if par is not None and par.tag != q("tabs"):
                out.append("\t")
        elif tag in (q("br"), q("cr")):
            out.append("\n")
        elif tag == q("noBreakHyphen"):
            out.append("-")
    return "".join(out)


def cell_text(tc) -> str:
    return "\n".join(para_text(p) for p in tc.findall(q("p")))


def has_content(p) -> bool:
    if para_text(p).strip():
        return True
    return any(True for _ in p.iter(q("drawing"), q("pict"), q("object")))


def para_atoms(p):
    """Последовательность 'c' (содержимое) и 'b' (разрыв страницы) в абзаце."""
    atoms = []
    ppr = p.find(q("pPr"))
    if ppr is not None and is_on(ppr.find(q("pageBreakBefore"))):
        atoms.append("b")
    for el in p.iter():
        tag = el.tag
        if tag == q("t") and (el.text or "").strip():
            atoms.append("c")
        elif tag == q("br") and wa(el, "type") == "page":
            atoms.append("b")
        elif tag in (q("drawing"), q("pict"), q("object")):
            atoms.append("c")
    if "c" not in atoms and ppr is not None and ppr.find(q("pBdr")) is not None:
        atoms.append("c")  # горизонтальная линия — тоже содержимое
    return atoms


# --------------------------------------------------------------------------
# Модель Markdown (свой разбор поверх markdown-it-py)
# --------------------------------------------------------------------------

MARKER_RE = re.compile(r"^<!--\s*(landscape|portrait|pagebreak)\s*-->$", re.I)
PAGEBREAK_LINES = {"\\newpage", "\\pagebreak"}
BR_RE = re.compile(r"^</?br\s*/?>$", re.I)  # </br> браузеры читают как <br>
FM_LINE = re.compile(r"^(\s|#|-\s|[^\s:#][^:]*:(\s|$))")  # §2: ключ, продолжение с отступом, «#», «- »
FM_KEY = re.compile(r"^[^\s:#][^:]*:(\s|$)")
XML_INVALID_RE = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")


VT_FF_ENTITY = re.compile(r"&#(?:0*1[12]|[xX]0*[bBcC]);")  # &#11; и &#12; в любой записи
SKIP_TAG_RE = re.compile(r"<(script|style|template)\b", re.I)
NEST_OPEN = ("blockquote_open", "bullet_list_open", "ordered_list_open", "list_item_open")


class MdInputError(ValueError):
    """md, который конвертер обязан отвергнуть (§2): docx по нему существовать не должен."""


@dataclass
class MdTable:
    """Таблица docx по md: таблица md или её часть (\\newpage между строками тела, §3.4)."""
    aligns: list
    rows: list  # строки → ячейки → текст (с \n на месте <br>)
    region: str = "P"  # ориентация участка (по маркерам/флагу)
    nested: bool = False  # внутри списка/цитаты
    headerless: bool = False  # все ячейки шапки пусты — строки шапки в docx нет (§3.4)
    cuts: list = field(default_factory=list)  # строки (0 — шапка), перед которыми \newpage
    group: int = 0  # номер таблицы md: у её частей общий
    part: int = 0  # номер части; у частей после первой шапки нет — все строки тела

    @property
    def head(self) -> bool:  # первая строка docx — шапка
        return not (self.headerless or self.part)


@dataclass
class MdList:
    id: int
    ordered: bool
    start: int
    delim: str
    depth: int
    in_quote: bool


@dataclass
class MdItem:
    list_id: int
    depth: int
    text: str


@dataclass
class MdBlock:
    """Блок верхнего уровня md — единица плана разделов (§4)."""
    kind: str  # heading, paragraph, hr, list, quote, code, table, html
    level: int = 0  # уровень заголовка
    b0: int = 0  # с какого куска проекции текста, картинки, таблицы md начинается блок
    i0: int = 0
    t0: int = 0
    words: int = 0
    chars: int = 0  # видимых символов абзаца (правило подтягивания, §4.3)
    colon: bool = False  # абзац оканчивается двоеточием — подпись к таблице
    image: bool = False
    hr: bool = False
    tables: list = field(default_factory=list)  # номера таблиц md внутри блока
    region: str = "P"  # ориентация участка
    start: bool = False  # серия маркеров перед блоком сменила ориентацию участка
    brk: bool = False  # в серии маркеров перед блоком был разрыв страницы
    parts: list = field(default_factory=list)  # куски проекции текста, с которых начинаются части
    breaks: list = field(default_factory=list)  # их номера слов в блоке: разрыв страницы (§3.4)

    @property
    def visible(self) -> bool:  # §4.1: есть текст, картинка, таблица или линия
        return bool(self.words or self.image or self.tables or self.hr)

    @property
    def barrier(self) -> bool:  # перед блоком новая страница: подтягивание выше не идёт
        return self.start or self.brk

    @property
    def simple(self) -> bool:  # короткий простой абзац (§4.3)
        return self.kind == "paragraph" and self.chars <= 600 and not self.image


@dataclass
class MdModel:
    path: Path
    text: str = ""
    headings: list = field(default_factory=list)  # (уровень, строки)
    tables: list = field(default_factory=list)
    images: list = field(default_factory=list)  # (alt, src)
    lists: list = field(default_factory=list)
    items: list = field(default_factory=list)
    title: str | None = None
    front_matter: bool = False
    html_blocks: int = 0
    flow: list = field(default_factory=list)  # блоки верхнего уровня (MdBlock) и маркеры (str) по порядку
    links: list = field(default_factory=list)  # адреса ссылок md по порядку
    too_deep: bool = False  # разбор оборвался по maxNesting
    code_blocks: list = field(default_factory=list)


def _no_footnote_def(state, start, end, silent):
    """§2: строка «[^метка]: текст» — не определение ссылки, а текст абзаца (сноски — текстом)."""
    if state.src.startswith("[^", state.bMarks[start] + state.tShift[start]):
        return False
    return _md_reference(state, start, end, silent)


def _vt_ff_entities(state):
    """§2: сущности &#11; и &#12; — пробел, а не склейка слов (до text_join, пока сущность видна)."""
    todo = [t.children for t in state.tokens if t.type == "inline"]
    while todo:
        for c in todo.pop() or []:
            if c.type == "text_special" and c.info == "entity" and VT_FF_ENTITY.fullmatch(c.markup):
                c.content = " "
            if c.children:
                todo.append(c.children)


def md_parser():
    # maxNesting 100 (SPEC §2): с пресетным 20 разборщик молча отбрасывает хвост после
    # списка глубины 10 — и сверщик не увидел бы потерю, разбирая md тем же способом
    md = MarkdownIt("commonmark", {"html": True, "maxNesting": 100}).enable(
        ["table", "strikethrough"])
    md.block.ruler.at("reference", _no_footnote_def)
    md.core.ruler.before("text_join", "vt_ff_entities", _vt_ff_entities)
    # §3.6: ссылка с любой схемой — ссылка (гиперссылкой станет только http/https/mailto/tel),
    # картинка `file:` — локальный файл; чёрный список markdown-it оставил бы их разметкой в тексте
    md.validateLink = lambda url: True
    return md


def strip_front_matter(text: str):
    lines = text.split("\n")
    if not lines or lines[0].strip() != "---" or len(lines) < 2 or not lines[1].strip():
        return text, False  # за --- пустая строка — линия, а не метаданные (как в pandoc)
    for j in range(1, min(len(lines), 50)):
        if lines[j].strip() in ("---", "..."):  # §2: закрывает и «...» (как в pandoc и YAML)
            body = lines[1:j]
            if any(FM_KEY.match(x) for x in body) and all(not x.strip() or FM_LINE.match(x) for x in body):
                return "\n".join(lines[j + 1:]), True
            return text, False
    return text, False


def _inline_text(children, images, src_dir: Path | None = None, links=None) -> str:
    parts, skip = [], None
    for c in children or []:
        t = c.type
        if skip:  # §2: содержимое inline <script>, <style>, <template> — не текст
            if t == "html_inline" and re.match(rf"</{skip}\s*>", c.content, re.I):
                skip = None
            continue
        if t in ("text", "text_special", "code_inline"):
            parts.append(c.content)
        elif t == "softbreak":
            parts.append(" ")
        elif t == "hardbreak":
            parts.append("\n")
        elif t == "html_inline":
            mm = SKIP_TAG_RE.match(c.content)
            if mm:
                skip = mm.group(1).lower()
            elif BR_RE.match(c.content.strip()):
                parts.append("\n")
        elif t == "link_open" and links is not None:
            links.append(c.attrGet("href") or "")
        elif t == "image":
            alt = _inline_text(c.children, [], src_dir)
            images.append((alt, c.attrGet("src") or ""))
    return "".join(parts)


BLOCK_TAG_RE = re.compile(r"</?(?:br|p|div|li|ul|ol|tr|td|th|h[1-6]|table|thead|tbody|"
                          r"blockquote|pre|hr|section|article|header|footer|details|summary|dl|dt|dd)\b[^>]*>",
                          re.I)


def _html_block_text(content: str, links=None) -> str:
    """Текст HTML-блока: блочные теги → перенос, прочие теги вырезаются без пробела
    (как в браузере: <b>HTML</b>-блока остаётся одним словом)."""
    s = re.sub(r"<!--.*?(?:-->|\Z)", " ", content, flags=re.S)  # незакрытый — до конца блока
    s = re.sub(r"<(script|style|template)\b.*?</\1\s*>", "", s, flags=re.S | re.I)  # не текст (§2)
    s = BLOCK_TAG_RE.sub("\n", s)
    s = re.sub(r"<[^>]*>", "", s)
    toks = md_parser().parseInline(s)  # §2: остаток — inline-Markdown, разметка в docx не попадает
    return _inline_text(toks[0].children if toks else [], [], links=links)


def is_comment_block(content: str) -> bool:
    return not re.sub(r"<!--.*?(?:-->|\Z)", "", content, flags=re.S).strip()


def _pagebreak_parts(children, content: str | None = None) -> list:
    """Inline-дети абзаца → части между строками \\newpage / \\pagebreak (SPEC §4.1: «строки»).
    Строка-маркер — текст верхнего уровня абзаца целиком на своей строке, в т. ч. посреди
    абзаца без пустых строк вокруг (как <!-- pagebreak -->, который markdown-it тоже выделяет
    из абзаца); регистр не важен. Текст «… \\newpage …» внутри строки — не маркер.
    content — сырой текст абзаца (inline-токен): разрыв — только строка, которая в исходнике
    дословно \\newpage; «\\\\newpage», «&#92;newpage» — текст, хотя после разбора выглядят так же
    (§4.1)."""
    kids = list(children or [])
    raw = content.split("\n") if content is not None else []
    parts, cur, depth, k, line = [], [], 0, 0, 0
    breaks = ("softbreak", "hardbreak")
    while k < len(kids):
        c = kids[k]
        cand = depth == 0 and c.type == "text" and c.content.strip().lower() in PAGEBREAK_LINES \
            and (k == 0 or kids[k - 1].type in breaks) and (k == len(kids) - 1 or kids[k + 1].type in breaks)
        if cand:  # строка кандидата — первая не выше счётчика, что даёт этот текст в любой записи
            # (счётчик отстаёт на `коде` в несколько строк: перевод строки в нём разбор не сохраняет)
            line = next((r for r in range(line, len(raw))
                         if unescapeAll(raw[r].strip().lower()) in PAGEBREAK_LINES), line)
        if cand and (line >= len(raw) or raw[line].strip().lower() in PAGEBREAK_LINES):
            if cur and cur[-1].type in breaks:
                cur.pop()
            parts.append(cur)
            cur, k, line = [], k + 2, line + 1  # перевод строки после маркера тоже уходит
            continue
        if c.type in breaks:
            line += 1
        elif c.type == "html_inline":
            line += c.content.count("\n")
        depth += c.nesting
        cur.append(c)
        k += 1
    parts.append(cur)
    return parts


def read_md(path: Path) -> str:
    """§2: байты → текст. BOM UTF-8 снимается, BOM UTF-16 LE/BE — декодируется; без BOM байт 0 —
    отказ (UTF-16 без BOM или двоичный файл); иначе строгий UTF-8."""
    data = Path(path).read_bytes()
    try:
        if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
            return data.decode("utf-16")
        if data.startswith(b"\xef\xbb\xbf"):
            return data[3:].decode("utf-8")
        if b"\x00" in data:
            raise MdInputError("байт 0 без BOM: похоже на UTF-16 без BOM или двоичный файл")
        return data.decode("utf-8")
    except UnicodeDecodeError as e:
        raise MdInputError(f"не декодируется как {e.encoding} (байт {e.start})") from None


def _table_break(line: str):
    """Строка таблицы — \\newpage / \\pagebreak (в списке и цитате — после отступа и «>»)?
    → отступ и «>» контейнера (пустая строка в нём — его строка без текста) или None."""
    lead = re.match(r"[ \t>]*", line).group()
    return lead if line[len(lead):].strip().lower() in PAGEBREAK_LINES else None


def split_table_breaks(text: str) -> str:
    """§4.1, §3.4: строка \\newpage / \\pagebreak под строкой таблицы (по GFM стала бы её строкой).
    Под ней (под серией таких строк) конец таблицы, новая таблица со своей шапкой или строка без
    «|» — разрыв после таблицы или между таблицами: строка отделяется пустой строкой. Под ней строка
    той же таблицы с «|» — строка остаётся строкой таблицы: по ней таблица режется на части
    (_parse_table)."""
    lines, found = text.split("\n"), {}  # строка-разрыв → конец её таблицы
    for tok in md_parser().parse(text):
        if tok.type == "table_open" and tok.map:
            for i in range(tok.map[0] + 2, tok.map[1]):
                if _table_break(lines[i]) is not None:
                    found[i] = tok.map[1]
    # где начались бы таблицы, если отделить пустой строкой (в том же контейнере) каждую
    # строку-разрыв
    src, trial = [], []  # номер строки исходника у каждой строки пробного текста (None — вставленная)
    for i, line in enumerate(lines):
        if i in found:
            src.append(None)
            trial.append(_table_break(line).rstrip())
        src.append(i)
        trial.append(line)
    trial = md_parser().parse("\n".join(trial))
    starts = {src[tok.map[0]] for tok in trial if tok.type == "table_open" and tok.map}
    cut = []
    for i, end in found.items():
        j = i + 1
        while j in found:  # серия строк-разрывов подряд
            j += 1
        if j >= end or j in starts or "|" not in lines[j]:
            cut.append(i)
    for i in sorted(cut, reverse=True):
        lines.insert(i, _table_break(lines[i]).rstrip())
    return "\n".join(lines)


BLOCK_KIND = {"heading_open": "heading", "paragraph_open": "paragraph", "hr": "hr",
              "fence": "code", "code_block": "code", "html_block": "html", "table_open": "table",
              "blockquote_open": "quote", "bullet_list_open": "list", "ordered_list_open": "list"}


def parse_md(path, landscape: bool = False) -> MdModel:
    path = Path(path)
    raw = read_md(path).replace("\r\n", "\n").replace("\r", "\n")
    raw = re.sub("[\x0b\x0c]", " ", raw)  # §2: VT и FF разделяют слова — пробел, а не склейка
    raw = XML_INVALID_RE.sub("", raw)  # SPEC §2: символы, недопустимые в XML, вырезаются
    text, fm = strip_front_matter(raw)
    text = split_table_breaks(text)
    tokens = md_parser().parse(text)
    lines = text.split("\n")
    m = MdModel(path=path, front_matter=fm)
    # §2: разбор оборвался по maxNesting — хвост глубже 100 уровней не попал в дерево
    m.too_deep = any(t.type in NEST_OPEN and t.level >= 99 for t in tokens)
    blocks = []
    list_stack, quote_depth = [], 0

    def block(kind, level=0):  # новый блок верхнего уровня в плане разделов
        b = MdBlock(kind, level, len(blocks), len(m.images), len(m.tables))
        m.flow.append(b)
        return b

    i = 0
    while i < len(tokens):
        tok = tokens[i]
        t = tok.type
        top = tok.level == 0 and t in BLOCK_KIND
        if top and t == "html_block" and is_comment_block(tok.content):
            mm = MARKER_RE.match(tok.content.strip())
            if mm:  # комментарий без маркера — не содержимое
                m.flow.append(mm.group(1).lower())
            i += 1
            continue
        if top and t != "paragraph_open":
            blk = block(BLOCK_KIND[t], int(tok.tag[1:]) if t == "heading_open" else 0)
            blk.hr = t == "hr"
        if t == "heading_open":
            s = _inline_text(tokens[i + 1].children, m.images, links=m.links)
            level = int(tok.tag[1:])
            m.headings.append((level, norm_lines(s)))
            if level == 1 and m.title is None:
                m.title = norm_ws(s)
            blocks.append(s)
            i += 3
            continue
        if t == "paragraph_open":
            # строки-маркеры разрыва вырезаются (вложенные в список/цитату — игнорируются)
            texts = []
            for k, part in enumerate(_pagebreak_parts(tokens[i + 1].children, tokens[i + 1].content)):
                if top and k:
                    m.flow.append("pagebreak")
                if part:
                    blk = block("paragraph") if top else None
                    texts.append(_inline_text(part, m.images, links=m.links))
                    blocks.append(texts[-1])
                    if blk:
                        blk.chars = len(norm_ws(texts[-1]))
                        blk.colon = norm_ws(texts[-1]).endswith(":")
            if i > 0 and tokens[i - 1].type == "list_item_open" and texts:
                m.items[-1].text = texts[0]  # номер стоит на первом абзаце пункта
            i += 3
            continue
        if t in ("fence", "code_block"):
            c = tok.content[:-1] if tok.content.endswith("\n") else tok.content
            blocks.append(c)
            m.code_blocks.append(c)
        elif t == "html_block":
            if not is_comment_block(tok.content):
                s = _html_block_text(tok.content, m.links)
                if s.strip():
                    blocks.append(s)
                    m.html_blocks += 1
        elif t == "table_open":
            tab, i = _parse_table(tokens, i, m.images, m.links, lines)
            tab.nested = tok.level > 0
            # §3.4: \newpage между строками тела — части таблицы с новой страницы, шапка только у
            # первой; над первой строкой тела — разрыв перед таблицей; внутри списка и цитаты разрыв
            # игнорируется (§4.1)
            cuts = [c for c in tab.cuts if c > 1] if top else []
            if top and 1 in tab.cuts:
                m.flow.insert(len(m.flow) - 1, "pagebreak")  # перед блоком таблицы
            group = m.tables[-1].group + 1 if m.tables else 0
            for k, (lo, hi) in enumerate(zip([0] + cuts, cuts + [len(tab.rows)])):
                if k:
                    blk.parts.append(len(blocks))
                part = MdTable(tab.aligns, tab.rows[lo:hi], nested=tab.nested,
                               headerless=tab.headerless and not k, group=group, part=k)
                m.tables.append(part)
                for row in part.rows:
                    blocks.extend(row)
        elif t in ("bullet_list_open", "ordered_list_open"):
            ordered = t == "ordered_list_open"
            start = to_int(tok.attrGet("start")) if ordered else None
            lst = MdList(len(m.lists), ordered, start if start is not None else 1,
                         tok.markup, len(list_stack), quote_depth > 0)
            m.lists.append(lst)
            list_stack.append(lst)
        elif t in ("bullet_list_close", "ordered_list_close"):
            list_stack.pop()
        elif t == "list_item_open":
            # каждый пункт md — ровно один абзац с номером; пустой пункт или пункт, который
            # начинается не с абзаца (код, вложенный список), — номер на пустом абзаце
            lst = list_stack[-1]
            m.items.append(MdItem(lst.id, lst.depth, ""))
        elif t == "blockquote_open":
            quote_depth += 1
        elif t == "blockquote_close":
            quote_depth -= 1
        i += 1
    m.text = "\n".join(blocks)
    _plan_regions(m, blocks, landscape)
    return m


def _plan_regions(m: MdModel, blocks: list, landscape: bool):
    """Объём блоков верхнего уровня и ориентация участков по маркерам (§4.1, §4.2).
    Серия маркеров подряд (невидимые блоки её не прерывают): ориентация — последний маркер
    ориентации, разрыв — если в серии был хоть один; новый участок и так с новой страницы."""
    blks = [x for x in m.flow if isinstance(x, MdBlock)]
    words = [len(x.split()) for x in blocks]
    ends = [(b.b0, b.i0, b.t0) for b in blks[1:]] + [(len(blocks), len(m.images), len(m.tables))]
    for b, (b1, i1, t1) in zip(blks, ends):
        b.words, b.image, b.tables = sum(words[b.b0:b1]), i1 > b.i0, list(range(b.t0, t1))
        b.breaks = [sum(words[b.b0:p]) for p in b.parts]
    region, series, seen = ("L" if landscape else "P"), [], False
    for x in m.flow:
        if isinstance(x, str):
            series.append(x)
            continue
        if x.visible:
            orients = [s for s in series if s != "pagebreak"]
            o = ("L" if orients[-1] == "landscape" else "P") if orients else None
            if not seen:
                region = o or region  # маркер до первого содержимого — без разрыва
            elif o and o != region:
                region, x.start = o, True
            else:
                x.brk = "pagebreak" in series
            series, seen = [], True
        x.region = region
        for k in x.tables:
            m.tables[k].region = region


def _parse_table(tokens, i, images, links=None, lines=()):
    """lines — строки разобранного текста: строка таблицы \\newpage / \\pagebreak (_table_break)
    — разрыв (split_table_breaks оставил его в таблице), а не строка."""
    aligns, rows, row, head_img, cuts = [], [], None, False, []
    j, n_img = i + 1, len(images)
    while tokens[j].type != "table_close":
        tt = tokens[j]
        if tt.type == "tr_open" and tt.map and tt.map[0] < len(lines) \
                and _table_break(lines[tt.map[0]]) is not None:
            if len(rows) not in cuts:
                cuts.append(len(rows))
            while tokens[j].type != "tr_close":
                j += 1
        elif tt.type == "tr_open":
            row = []
        elif tt.type == "tr_close":
            rows.append(row)
            if len(rows) == 1:
                head_img = len(images) > n_img
        elif tt.type in ("th_open", "td_open"):
            mm = re.search(r"text-align:\s*(\w+)", tt.attrGet("style") or "")
            if tt.type == "th_open":
                aligns.append(mm.group(1) if mm else "")
            nxt = tokens[j + 1]
            row.append(_inline_text(nxt.children, images, links=links) if nxt.type == "inline" else "")
        j += 1
    # §3.4: шапка без видимого текста и картинок в docx не выводится
    headerless = len(rows) > 1 and not head_img and not any(norm_ws(c) for c in rows[0])
    return MdTable(aligns, rows, headerless=headerless, cuts=cuts), j


# --------------------------------------------------------------------------
# Проверки: §5 порядок и структура
# --------------------------------------------------------------------------


def _check_children_order(rep, el, rank, where, kind, code="order"):
    last_rank, last_name, seen = -1, None, set()
    for ch in el:
        name = local(ch)
        if name is None:
            continue  # комментарии и чужие пространства (w14 и т. п. — mc:Ignorable)
        if name not in rank:
            rep.err(code, f"{where}: в <w:{kind}> неизвестный элемент <w:{name}>")
            continue
        r = rank[name]
        if r < last_rank:
            rep.err(code, f"{where}: в <w:{kind}> элемент <w:{name}> стоит после "
                          f"<w:{last_name}> — нарушен порядок схемы")
        if name in seen and name not in REPEATABLE:
            rep.err(code, f"{where}: в <w:{kind}> элемент <w:{name}> повторяется")
        seen.add(name)
        if r >= last_rank:
            last_rank, last_name = r, name


def _where(el, part):
    """Короткая привязка элемента к месту в документе."""
    anc = el
    while anc is not None and local(anc) not in ("p", "tbl", "style", "sectPr", "lvl",
                                                    "abstractNum", "num", "docDefaults", "font"):
        anc = anc.getparent()
    line = el.sourceline
    tag = local(anc) if anc is not None else local(el)
    sid = wa(anc, "styleId") if anc is not None and local(anc) == "style" else None
    return f"{part}:{line} ({'стиль ' + sid if sid else tag})"


def check_order(rep, dx: Docx):
    count = 0
    for part, root in (("document.xml", dx.document), ("styles.xml", dx.styles),
                       ("fontTable.xml", dx.font_table)):
        if root is None:
            continue
        for el in root.iter():
            name = local(el)
            if name in ORDER_RANK:
                count += 1
                _check_children_order(rep, el, ORDER_RANK[name], _where(el, part), name)
    if dx.numbering is not None:
        root = dx.numbering
        for el in root.iter():
            name = local(el)
            if el is root:
                _check_children_order(rep, el, NUMBERING_RANK["numbering"], "numbering.xml",
                                      "numbering", code="order")
            elif name in NUMBERING_RANK:
                _check_children_order(rep, el, NUMBERING_RANK[name], _where(el, "numbering.xml"),
                                      name, code="order")
            elif name in ("pPr", "rPr"):
                _check_children_order(rep, el, ORDER_RANK[name], _where(el, "numbering.xml"), name)
    rep.info.append(f"Порядок §5: проверено элементов {count} (+ numbering.xml)")


def check_structure(rep, dx: Docx):
    body = dx.body
    kids = [c for c in body if local(c)]
    if not kids or local(kids[-1]) != "sectPr":
        rep.err("structure", "последним элементом w:body должен быть w:sectPr")
    for p in body.iter(q("p")):
        ks = [c for c in p if local(c)]
        if any(local(c) == "pPr" for c in ks[1:]):
            rep.err("structure", f"document.xml:{p.sourceline}: w:pPr не первым в абзаце")
    for r in body.iter(q("r")):
        ks = [c for c in r if local(c)]
        if any(local(c) == "rPr" for c in ks[1:]):
            rep.err("structure", f"document.xml:{r.sourceline}: w:rPr не первым в прогоне")
    for tbl in body.iter(q("tbl")):
        ks = [local(c) for c in tbl if local(c)]
        if ks[:2] != ["tblPr", "tblGrid"]:
            rep.err("structure", f"document.xml:{tbl.sourceline}: таблица должна начинаться "
                                 f"с tblPr, tblGrid, а начинается с {ks[:2]}")
        for tr in tbl.findall(q("tr")):
            rk = [local(c) for c in tr if local(c)]
            first_tc = rk.index("tc") if "tc" in rk else len(rk)
            head = rk[:first_tc]
            if any(x not in ("tblPrEx", "trPr") for x in head) or head not in (
                    [], ["trPr"], ["tblPrEx"], ["tblPrEx", "trPr"]):
                rep.err("structure", f"document.xml:{tr.sourceline}: в строке таблицы перед "
                                     f"ячейками {head}")
            if any(x in ("trPr", "tblPrEx") for x in rk[first_tc:]):
                rep.err("structure", f"document.xml:{tr.sourceline}: trPr после ячеек")
            for tc in tr.findall(q("tc")):
                ck = [local(c) for c in tc if local(c)]
                if "tcPr" in ck[1:]:
                    rep.err("structure", f"document.xml:{tc.sourceline}: tcPr не первым в ячейке")
                blocks = [x for x in ck if x in ("p", "tbl", "sdt", "customXml")]
                if not blocks or blocks[-1] != "p":
                    rep.err("structure", f"document.xml:{tc.sourceline}: ячейка должна "
                                         f"заканчиваться абзацем w:p (Word иначе считает файл "
                                         f"повреждённым)")
    for hl in body.iter(q("hyperlink")):
        rid = hl.get("{%s}id" % R_NS)
        if rid is None:
            if wa(hl, "anchor") is None:
                rep.err("structure", f"document.xml:{hl.sourceline}: w:hyperlink без r:id и w:anchor")
            continue
        rel = dx.rel_by_id.get(rid)
        if rel is None:
            rep.err("structure", f"гиперссылка ссылается на несуществующую связь {rid}")
        elif rel.get("TargetMode") != "External":
            rep.err("structure", f"гиперссылка {rid} → {rel.get('Target')}: связь не External")
    for blip in body.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
        rid = blip.get("{%s}embed" % R_NS)
        if rid not in dx.rel_by_id:
            rep.err("structure", f"картинка ссылается на несуществующую связь {rid}")


CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
# последний сегмент типа связи → окончание (или начало, если с «/») типа содержимого части
REL_CONTENT_TYPE = {
    "officeDocument": ".main+xml",
    "styles": "wordprocessingml.styles+xml",
    "numbering": "wordprocessingml.numbering+xml",
    "settings": "wordprocessingml.settings+xml",
    "fontTable": "wordprocessingml.fontTable+xml",
    "webSettings": "wordprocessingml.webSettings+xml",
    "footnotes": "wordprocessingml.footnotes+xml",
    "endnotes": "wordprocessingml.endnotes+xml",
    "header": "wordprocessingml.header+xml",
    "footer": "wordprocessingml.footer+xml",
    "comments": "wordprocessingml.comments+xml",
    "theme": "officedocument.theme+xml",
    "core-properties": "package.core-properties+xml",
    "extended-properties": "officedocument.extended-properties+xml",
    "customXmlProps": "customXmlProperties+xml",
    "stylesWithEffects": "stylesWithEffects+xml",
    "image": "image/",
    "thumbnail": "image/",
}


def check_package(rep, dx: Docx):
    """Целостность пакета — то, на что Word отвечает «файл повреждён» (SPEC §6(в)):
    у каждой части есть тип содержимого, внутренние связи ведут в существующие части,
    префиксы из mc:Ignorable объявлены, id рисунков (wp:docPr) уникальны."""
    from posixpath import dirname, join, normpath
    from urllib.parse import unquote

    names = {n for n in dx.parts if not n.endswith("/")}
    lower = {n.lower(): n for n in names}
    ct_raw = dx.parts.get("[Content_Types].xml")
    if ct_raw is None:
        rep.err("package", "нет [Content_Types].xml")
        return
    xml = {}
    for n in sorted(names):
        if n.lower().endswith((".xml", ".rels")):
            try:
                xml[n] = etree.fromstring(dx.parts[n])
            except etree.XMLSyntaxError as e:
                rep.err("package", f"{n}: XML не разбирается — {e}")
    ct = xml.get("[Content_Types].xml")
    if ct is None:
        return
    defaults = {(d.get("Extension") or "").lower(): d.get("ContentType") or ""
                for d in ct.findall("{%s}Default" % CT_NS)}
    overrides = {(o.get("PartName") or "").lstrip("/").lower(): o.get("ContentType") or ""
                 for o in ct.findall("{%s}Override" % CT_NS)}

    def content_type(n):
        ext = n.rsplit(".", 1)[-1].lower() if "." in n.rsplit("/", 1)[-1] else ""
        return overrides.get(n.lower(), defaults.get(ext))

    for n in sorted(names - {"[Content_Types].xml"}):
        if content_type(n) is None:
            rep.err("package", f"{n}: нет типа содержимого в [Content_Types].xml")
    for o in sorted(set(overrides) - set(lower)):
        rep.warn("package", f"[Content_Types].xml: Override для отсутствующей части /{o}")
    for n, root in sorted(xml.items()):
        if n.endswith(".rels"):
            src_dir = dirname(dirname(n))  # word/_rels/document.xml.rels → word
            ids = Counter(r.get("Id") for r in root)
            for rid, k in ids.items():
                if k > 1:
                    rep.err("package", f"{n}: Id связи {rid} повторяется {k} раз")
            for r in root:
                if not isinstance(r.tag, str) or r.get("TargetMode") == "External":
                    continue
                tgt = unquote(r.get("Target") or "")
                path = normpath(tgt.lstrip("/") if tgt.startswith("/") else join(src_dir, tgt))
                if path.lower() not in lower:
                    rep.err("package", f"{n}: связь {r.get('Id')} ведёт в отсутствующую часть "
                                       f"{path}")
                    continue
                # тип части должен соответствовать типу связи: при запасном Default
                # «xml → application/xml» Word не узнаёт стили/нумерацию/настройки
                kind = (r.get("Type") or "").rsplit("/", 1)[-1]
                got = content_type(lower[path.lower()]) or ""
                want = REL_CONTENT_TYPE.get(kind)
                if want and not (got.startswith(want) if want.endswith("/") else got.endswith(want)):
                    rep.err("package", f"{path}: тип содержимого «{got}» не подходит для связи "
                                       f"{kind} (нужен …{want})")
        ign = root.get("{%s}Ignorable" % MC_NS)
        if ign:
            undeclared = [p for p in ign.split() if p not in root.nsmap]
            if undeclared:
                rep.err("package", f"{n}: в mc:Ignorable префиксы без объявления пространства "
                                   f"имён {undeclared} — Word считает файл повреждённым")
    ids = Counter(e.get("id") for e in dx.document.iter("{%s}docPr" % WP_NS))
    for i, k in ids.items():
        if k > 1:
            rep.err("package", f"document.xml: id рисунка (wp:docPr) {i} повторяется {k} раз")


# --------------------------------------------------------------------------
# §3.1 страница, §4.4 разделы и разрывы
# --------------------------------------------------------------------------


def parse_orient_list(s: str):
    out = []
    for x in re.split(r"[,\s]+", s.strip()):
        if not x:
            continue
        o = ORIENT_ALIASES.get(x.upper())
        if o is None:
            raise ValueError(f"неизвестная ориентация «{x}» (ожидается P/L или К/А)")
        out.append(o)
    return out


def check_sections(rep, dx: Docx, expect=None, empty_md=False):
    secs = dx.sections
    orients = [s.orient for s in secs]
    ru = ", ".join("К" if o == "P" else "А" if o == "L" else "?" for o in orients)
    rep.info.append(f"Разделы: {len(secs)} ({ru})")
    if expect is not None and orients != expect:
        exp_ru = ", ".join("К" if o == "P" else "А" for o in expect)
        rep.err("orient", f"ориентация разделов {ru} — ожидалось {exp_ru}")
    for s in secs:
        lab = f"раздел {s.index + 1}"
        sp = s.sectPr
        if sp is None:
            rep.err("page", f"{lab}: нет w:sectPr у документа (последний раздел не задан)")
            continue
        pg, mar = sp.find(q("pgSz")), sp.find(q("pgMar"))
        if pg is None:
            rep.err("page", f"{lab}: нет pgSz — размер страницы не задан явно")
        else:
            w_, h_ = to_int(wa(pg, "w")), to_int(wa(pg, "h"))
            o = wa(pg, "orient")
            exp = A4["L" if (o == "landscape" or (w_ or 0) > (h_ or 0)) else "P"]
            if (w_, h_) != exp:
                rep.err("page", f"{lab}: pgSz {w_}×{h_}, ожидалось A4 {exp[0]}×{exp[1]}")
            if o == "landscape" and w_ is not None and h_ is not None and w_ <= h_:
                rep.err("page", f"{lab}: orient=landscape, но ширина {w_} ≤ высоты {h_}")
            if (w_ or 0) > (h_ or 0) and o != "landscape":
                rep.err("page", f"{lab}: страница шире высоты, но нет w:orient=\"landscape\"")
            if o == "portrait" and (w_ or 0) > (h_ or 0):
                rep.err("page", f"{lab}: orient=portrait при ширине больше высоты")
        if mar is None:
            rep.err("page", f"{lab}: нет pgMar — поля не заданы явно")
        else:
            bad = []
            for side in ("top", "right", "bottom", "left"):
                v = to_int(wa(mar, side))
                if v != MARGIN:
                    bad.append(f"{side}={v}")
            for side in ("header", "footer"):
                v = to_int(wa(mar, side))
                if v != HDR_FTR:
                    bad.append(f"{side}={v}")
            g = to_int(wa(mar, "gutter"))
            if g not in (None, 0):
                bad.append(f"gutter={g}")
            if bad:
                rep.err("page", f"{lab}: поля {', '.join(bad)} — ожидалось 1134 со всех "
                                f"сторон, header/footer 720")
        for ref in ("headerReference", "footerReference"):
            if sp.find(q(ref)) is not None:
                rep.err("page", f"{lab}: есть {ref} — колонтитулов быть не должно")
        typ = wa(sp.find(q("type")), "val")
        if typ not in (None, "nextPage"):
            rep.err("sections", f"{lab}: тип разрыва {typ}, ожидался nextPage")
        tw = s.text_width
        if tw is not None and s.orient in TEXT_WIDTH and tw != TEXT_WIDTH[s.orient]:
            rep.err("page", f"{lab}: ширина текста {tw}, ожидалось {TEXT_WIDTH[s.orient]}")
    # пустые разделы, разрывы страниц
    for s in secs:
        lab = f"раздел {s.index + 1}"
        atoms = []
        for el in s.elements:
            if el.tag == q("tbl"):
                atoms.append("c")
            elif el.tag == q("p"):
                atoms.extend(para_atoms(el))
            elif el.tag == q("sdt"):
                if any(True for _ in el.iter(q("t"))):
                    atoms.append("c")
        if "c" not in atoms and not (empty_md and len(secs) == 1):  # пустой md — одна пустая страница
            rep.err("sections", f"{lab} пуст — нет ни текста, ни таблицы (пустая страница)")
            continue
        prev = None
        for k, a in enumerate(atoms):
            if a == "b":
                if "c" not in atoms[:k]:
                    rep.err("sections", f"{lab}: разрыв страницы до первого содержимого "
                                        f"раздела — пустая страница")
                elif prev == "b":
                    rep.err("sections", f"{lab}: два разрыва страницы подряд — пустая страница")
            prev = a
        if atoms and atoms[-1] == "b":
            rep.err("sections", f"{lab}: разрыв страницы после последнего содержимого "
                                f"раздела — пустая страница")
        car = s.carrier
        if car is not None and not has_content(car):
            sp_ = car.find("w:pPr/w:spacing", NS)
            vals = {a: wa(sp_, a) for a in ("before", "after", "line", "lineRule")}
            if (to_int(vals["before"]) or 0) != 0 or to_int(vals["after"]) != 0 or \
                    to_int(vals["line"]) != 20 or vals["lineRule"] != "exact":
                rep.err("sections", f"{lab}: абзац-носитель разрыва не обнулён: spacing "
                                    f"{vals}, ожидалось before=0 after=0 line=20 exact")
            sz = to_int(wa(car.find("w:pPr/w:rPr/w:sz", NS), "val"))
            if sz != 2:
                rep.err("sections", f"{lab}: знак абзаца-носителя разрыва sz={sz}, ожидалось 2")


# --------------------------------------------------------------------------
# §3.2 docDefaults, settings, fontTable, свойства документа
# --------------------------------------------------------------------------


def check_defaults(rep, dx: Docx):
    if dx.styles is None:
        rep.err("defaults", "нет word/styles.xml")
        return
    rpr, ppr = dx.rpr_default, dx.ppr_default
    if rpr is None:
        rep.err("defaults", "нет docDefaults/rPrDefault/rPr")
    else:
        rf = rpr.find(q("rFonts"))
        exp = {"ascii": HN, "hAnsi": HN, "cs": HN, "eastAsia": "Times New Roman"}
        got = {k: wa(rf, k) for k in exp}
        if got != exp:
            rep.err("defaults", f"docDefaults rFonts {got}, ожидалось {exp}")
        themes = [a for a in THEME_FONT_ATTRS if wa(rf, a) is not None]
        if themes:
            rep.err("defaults", f"docDefaults rFonts содержит theme-атрибуты {themes} — "
                                f"они перебивают явный шрифт")
        for tag in ("sz", "szCs"):
            v = to_int(wa(rpr.find(q(tag)), "val"))
            if v != 22:
                rep.err("defaults", f"docDefaults {tag}={v}, ожидалось 22 (11 pt)")
        lang = rpr.find(q("lang"))
        got = {k: wa(lang, k) for k in ("val", "eastAsia", "bidi")}
        if got != {"val": "ru-RU", "eastAsia": "ru-RU", "bidi": "ar-SA"}:
            rep.err("defaults", f"docDefaults lang {got}, ожидалось ru-RU/ru-RU/ar-SA")
    if ppr is None:
        rep.err("defaults", "нет docDefaults/pPrDefault/pPr")
    else:
        sp = ppr.find(q("spacing"))
        got = {k: wa(sp, k) for k in ("before", "after", "line", "lineRule")}
        if (to_int(got["before"]) or 0) != 0 or to_int(got["after"]) != 160 or \
                to_int(got["line"]) != 278 or got["lineRule"] not in ("auto", None):
            rep.err("defaults", f"docDefaults spacing {got}, ожидалось after=160 line=278 auto")
        if not is_on(ppr.find(q("widowControl"))):
            rep.err("defaults", "docDefaults: нет widowControl")
        jc = wa(ppr.find(q("jc")), "val")
        if jc not in (None, "left", "start"):
            rep.err("defaults", f"docDefaults jc={jc}, ожидалось выравнивание влево")
    normal = dx.styles_by_id.get(dx.default_pstyle) if dx.default_pstyle else None
    if normal is None:
        rep.err("defaults", "нет стиля абзаца по умолчанию (Normal)")
    else:
        own = []
        for part in ("pPr", "rPr"):
            el = normal.find(q(part))
            if el is not None:
                own += [f"{part}/{local(c)}" for c in el if local(c)]
        if own:
            rep.err("defaults", f"стиль Normal переопределяет {own} — должен быть без "
                                f"собственных переопределений")
    # выравнивание по ширине нигде не допускается
    for s in dx.styles_by_id.values():
        jc = wa(s.find("w:pPr/w:jc", NS), "val")
        if jc in JUSTIFY:
            rep.err("defaults", f"стиль {wa(s, 'styleId')}: jc={jc} — по ширине выравнивать нельзя")
    nj = sum(1 for jc in dx.body.iter(q("jc")) if wa(jc, "val") in JUSTIFY
             and jc.getparent() is not None and local(jc.getparent()) == "pPr")
    if nj:
        rep.err("defaults", f"{nj} абзац(ев) выровнены по ширине — запрещено §3.2")
    # settings.xml
    if dx.settings is None:
        rep.err("defaults", "нет word/settings.xml")
    else:
        cm = None
        for cs in dx.settings.iter(q("compatSetting")):
            if wa(cs, "name") == "compatibilityMode":
                cm = wa(cs, "val")
        if cm != "15":
            rep.err("defaults", f"settings.xml compatibilityMode={cm}, ожидалось 15 "
                                f"(иначе Word рисует таблицу со сдвигом и «Режим совместимости»)")


def _used_fonts(dx: Docx):
    used = set()
    for root in (dx.document, dx.styles, dx.numbering):
        if root is None:
            continue
        for rf in root.iter(q("rFonts")):
            for a in ("ascii", "hAnsi", "cs", "eastAsia"):
                v = wa(rf, a)
                if v:
                    used.add(v)
    return used


def check_fonts(rep, dx: Docx):
    if dx.font_table is None:
        rep.err("fonts", "нет word/fontTable.xml — Word на Windows не узнает, чем заменить "
                         "Helvetica Neue")
        return
    fonts = {wa(f, "name"): f for f in dx.font_table.findall(q("font"))}
    used = _used_fonts(dx)

    def check(name, alt, family, pitch, panose=None, required=True):
        f = fonts.get(name)
        if f is None:
            if required:
                rep.err("fonts", f"в fontTable нет записи «{name}»")
            return
        got = {"altName": wa(f.find(q("altName")), "val"),
               "family": wa(f.find(q("family")), "val"),
               "pitch": wa(f.find(q("pitch")), "val")}
        exp = {"altName": alt, "family": family, "pitch": pitch}
        if panose:
            got["panose1"] = (wa(f.find(q("panose1")), "val") or "").upper() or None
            exp["panose1"] = panose
        bad = {k: got[k] for k in exp if got[k] != exp[k]}
        if bad:
            rep.err("fonts", f"fontTable «{name}»: {bad}, ожидалось "
                             f"{ {k: exp[k] for k in bad} }")

    check(HN, "Arial", "swiss", "variable", HN_PANOSE, required=True)
    check("Menlo", "Consolas", "modern", "fixed", required="Menlo" in used)
    missing = sorted(u for u in used if u not in fonts and not u.startswith("theme:"))
    if missing:
        rep.warn("fonts", f"шрифты используются, но не описаны в fontTable: {missing}")


CORE_FIELDS = (("title", DC_NS, "title"), ("author", DC_NS, "creator"),
               ("last_modified_by", CP_NS, "lastModifiedBy"),
               ("comments", DC_NS, "description"), ("subject", DC_NS, "subject"),
               ("keywords", CP_NS, "keywords"), ("revision", CP_NS, "revision"))


def core_props(dx: Docx) -> dict:
    out = {}
    if dx.core is None:
        return out
    for key, ns, tag in CORE_FIELDS:
        el = dx.core.find("{%s}%s" % (ns, tag))
        out[key] = (el.text or "") if el is not None else None
    return out


def check_core(rep, dx: Docx, expected_title=None):
    if dx.core is None:
        rep.err("core", "нет docProps/core.xml")
        return
    raw = etree.tostring(dx.core, encoding="unicode")
    if "python-docx" in raw:
        rep.err("core", "в свойствах документа осталось «python-docx»")
    if dx.app is not None and "python-docx" in etree.tostring(dx.app, encoding="unicode"):
        rep.warn("core", "в docProps/app.xml есть «python-docx»")
    cp = core_props(dx)
    for k in ("author", "last_modified_by", "comments", "subject", "keywords"):
        if cp.get(k):
            rep.err("core", f"свойство {k}=«{cp[k]}», ожидалось пустое")
    if (cp.get("revision") or "").strip() != "1":
        rep.err("core", f"revision={cp.get('revision')!r}, ожидалось 1")
    ident = dx.core.find(f"{{{DC_NS}}}identifier")
    if ident is None or (ident.text or "").strip() != "md2docx":
        rep.err("core", f"dc:identifier=«{'' if ident is None else ident.text}», ожидалось «md2docx» "
                        f"(метка «наш файл» для защиты от перезаписи, §1)")
    title = cp.get("title") or ""
    if len(title) > 255:
        rep.err("core", f"title длиной {len(title)} — больше 255 символов (§3.2)")
    if expected_title is not None:
        want = norm_ws(expected_title)
        if len(want) > 255:  # §3.2: длинный H1 — заголовок документа обрезан
            cut = norm_ws(title).rstrip("…").rstrip()
            ok = len(title) <= 255 and bool(cut) and want.startswith(cut)
        else:
            ok = norm_ws(title) == want
        if not ok:
            rep.err("core", f"title=«{short(title, 80)}», ожидалось «{short(expected_title, 80)}»")


# --------------------------------------------------------------------------
# §3.3 стили заголовков и абзацы заголовков
# --------------------------------------------------------------------------


def check_heading_styles(rep, dx: Docx):
    for n in range(1, 7):
        sid = f"Heading{n}"
        st = dx.styles_by_id.get(sid)
        if st is None:
            alt = dx.find_style_by_name(f"heading {n}")
            if alt is not None:
                rep.err("heading-styles", f"нет стиля styleId «{sid}»: заголовок {n} "
                                          f"оформлен стилем «{wa(alt, 'styleId')}»")
            else:
                rep.err("heading-styles", f"нет стиля «{sid}» (heading {n})")
            st = alt
            if st is None:
                continue
        lab = f"стиль {wa(st, 'styleId')} (H{n})"
        name = (wa(st.find(q("name")), "val") or "").lower()
        if name != f"heading {n}":
            rep.err("heading-styles", f"{lab}: имя «{name}», ожидалось «heading {n}» "
                                      f"(иначе не работают навигация и оглавление)")
        chain = [st] + dx.style_chain(wa(st.find(q("basedOn")), "val"))

        def p_el(child):
            for s in chain:
                el = s.find(f"w:pPr/w:{child}", NS)
                if el is not None:
                    return el
            return dx.ppr_default.find(q(child)) if dx.ppr_default is not None else None

        def p_attr(child, attr):
            for s in chain:
                v = wa(s.find(f"w:pPr/w:{child}", NS), attr)
                if v is not None:
                    return v
            return wa(dx.ppr_default.find(q(child)), attr) if dx.ppr_default is not None else None

        def r_el(child):
            for s in chain:
                el = s.find(f"w:rPr/w:{child}", NS)
                if el is not None:
                    return el
            return dx.rpr_default.find(q(child)) if dx.rpr_default is not None else None

        bad = []
        for flag in ("keepNext", "keepLines"):
            if not is_on(p_el(flag)):
                bad.append(f"нет {flag}")
        sz_exp, before_exp = HEADING_SPEC[n]
        b4, af = to_int(p_attr("spacing", "before")), to_int(p_attr("spacing", "after"))
        if b4 != before_exp:
            bad.append(f"before={b4} (нужно {before_exp})")
        if af != 160:
            bad.append(f"after={af} (нужно 160)")
        ol = to_int(wa(p_el("outlineLvl"), "val"))
        if ol != n - 1:
            bad.append(f"outlineLvl={ol} (нужно {n - 1})")
        if p_el("numPr") is not None:
            bad.append("есть numPr — автонумерации заголовков быть не должно")
        rsources = [s.find(q("rPr")) for s in chain] + [dx.rpr_default]
        for a in ("ascii", "hAnsi"):
            f = font_from_sources(rsources, a)
            if f != HN:
                bad.append(f"шрифт {a}={f} (нужен явный {HN})")
        own_rf = st.find("w:rPr/w:rFonts", NS)
        themes = [a for a in THEME_FONT_ATTRS if wa(own_rf, a) is not None]
        if themes:
            bad.append(f"theme-атрибуты шрифта {themes}")
        if not is_on(r_el("b")):
            bad.append("не полужирный")
        sz = to_int(wa(r_el("sz"), "val"))
        if sz != sz_exp:
            bad.append(f"sz={sz} (нужно {sz_exp})")
        if any(s.find("w:rPr/w:color", NS) is not None for s in chain):
            bad.append("задан w:color (нужен авто)")
        if n == 6 and not is_on(r_el("i")):
            bad.append("H6 не курсив")
        if bad:
            rep.err("heading-styles", f"{lab}: " + "; ".join(bad))


def check_heading_paragraphs(rep, dx: Docx):
    for p in dx.body.iter(q("p")):
        lvl = dx.heading_level(p)
        if lvl is None:
            continue
        txt = short(norm_ws(para_text(p)))
        sid = dx.pstyle_id(p)
        st_sz = None
        for s in dx.style_chain(sid):
            st_sz = to_int(wa(s.find("w:rPr/w:sz", NS), "val"))
            if st_sz is not None:
                break
        if p.find("w:pPr/w:numPr", NS) is not None:
            rep.err("heading-styles", f"заголовок «{txt}»: есть numPr — автонумерации быть не должно")
        unbold = resized = 0
        for r in p.iter(q("r")):
            if not (para_text(r)).strip():
                continue
            rpr = r.find(q("rPr"))
            if rpr is None:
                continue
            b = rpr.find(q("b"))
            if b is not None and not is_on(b):
                unbold += 1
            sz = to_int(wa(rpr.find(q("sz")), "val"))
            font = wa(rpr.find(q("rFonts")), "ascii")
            if sz is not None and sz != st_sz and font in CODE_FONTS:  # §3.6: код — кегль заголовка
                rep.err("heading-styles", f"заголовок «{txt}»: inline-код «{short(para_text(r), 20)}» "
                                          f"{sz / 2:g} pt, нужен кегль заголовка {(st_sz or 0) / 2:g} pt")
            elif sz is not None and sz != st_sz:
                resized += 1
        if unbold or resized:
            what = []
            if unbold:
                what.append(f"снят полужирный (w:b=0) в {unbold} прогон(ах)")
            if resized:
                what.append(f"размер шрифта перебит в {resized} прогон(ах)")
            rep.err("heading-styles", f"абзац со стилем заголовка H{lvl} «{txt}»: "
                                      f"{', '.join(what)} — обычный абзац, оформленный стилем "
                                      f"заголовка (в md заголовок не может быть нежирным)")


# --------------------------------------------------------------------------
# §3.4 таблицы, §4.3 широкие таблицы
# --------------------------------------------------------------------------


def run_style(dx: Docx, r, p, header: bool):
    bold = header or is_on(dx.eff_rpr(r, p, "b"))
    italic = is_on(dx.eff_rpr(r, p, "i"))
    font = dx.eff_font(r, p, "ascii") or ""
    code = font in CODE_FONTS
    sz = to_int(wa(dx.eff_rpr(r, p, "sz"), "val")) or 22
    return (code, bold, italic, sz / 2)


SPLIT_WS = re.compile(r"[ \t\n\r\f\v\u2002\u2003\u2009\u200b]+")  # §3.4: и узкие пробелы, ZWSP
DASH_SPLIT = re.compile(r"(?<=[-‐–—])")
URL_RE = re.compile(r"://|\bwww\.", re.I)


def cell_chars(dx: Docx, tc, header: bool):
    """Символы ячейки со стилем; граница абзаца, w:br и таб — (символ, None)."""
    seq = []  # (символ, стиль)
    for pi, p in enumerate(tc.findall(q("p"))):
        if pi:
            seq.append(("\n", None))
        for r in p.iter(q("r")):
            st = run_style(dx, r, p, header)
            for ch in r:
                if ch.tag == q("t"):
                    seq.extend((c, st) for c in (ch.text or ""))
                elif ch.tag in (q("br"), q("cr")):
                    seq.append(("\n", None))
                elif ch.tag == q("tab"):
                    seq.append(("\t", None))
    return seq


def piece_width(piece, meas: Measurer) -> float:
    """Ширина [(символ, стиль)] в twips: замер по участкам одного стиля."""
    width, seg, seg_st = 0.0, "", None
    for x, s in piece:
        if s != seg_st and seg:
            width += meas.twips(seg, seg_st[3], seg_st[0], seg_st[1], seg_st[2])
            seg = ""
        seg_st = s
        seg += x
    if seg:
        width += meas.twips(seg, seg_st[3], seg_st[0], seg_st[1], seg_st[2])
    return width


def cell_height(dx: Docx, tc, header: bool, meas: Measurer, avail: float) -> float:
    """Оценка высоты ячейки в twips (SPEC §3.4 «Строки»): жадный перенос слов по ширине
    колонки без полей, слово шире ячейки — посимвольно; строка ROW_LINE, поля 40 + 40; картинка
    добавляет высоту рисунка (wp:extent cy)."""
    avail, lines, x, word = max(avail, 1), 1, 0.0, []
    for c, st in cell_chars(dx, tc, header) + [(" ", None)]:
        if st is not None and not SPLIT_WS.fullmatch(c):
            word.append((c, st))
            continue
        if word:
            w = piece_width(word + [(" ", word[-1][1])], meas)
            if x + w > avail:
                lines += (1 if x else 0) + int(w // avail)
                x = w % avail
            else:
                x += w
            word = []
        if c == "\n":
            lines, x = lines + 1, 0.0
    pics = sum((to_int(e.get("cy")) or 0) / 635 for e in tc.iter(f"{{{WP_NS}}}extent"))
    return lines * ROW_LINE + 80 + pics


def cell_min_word(dx: Docx, tc, header: bool, meas: Measurer, soft=False):
    """(ширина самого длинного неразрывного слова в twips, само слово). Жёсткий min (§3.4):
    перенос после ‐ – — всегда, после «-» — кроме «не цифра, дефис, цифра» (Ф-123456);
    мягкий (soft) — по дефисам не делится. URL по дефисам не делится; кусок длиннее 30 символов
    не учитывается."""
    words, cur = [], []
    for c, st in cell_chars(dx, tc, header):
        if st is None or SPLIT_WS.fullmatch(c):
            if cur:
                words.append(cur)
            cur = []
        else:
            cur.append((c, st))
    if cur:
        words.append(cur)
    best, best_word = 0.0, ""
    for wd in words:
        piece = []
        url = URL_RE.search("".join(c for c, _ in wd))  # URL по дефисам не делится (§3.4)
        for k, (c, st) in enumerate(wd):
            piece.append((c, st))
            nc = wd[k + 1][0] if k + 1 < len(wd) else ""
            brk = c in "‐–—" or c == "-" and (not nc.isdecimal() or k > 0 and wd[k - 1][0].isdecimal())
            if (brk and not url and not soft) or k == len(wd) - 1:
                text = "".join(x for x, _ in piece)
                if len(text) <= 30:
                    width = piece_width(piece, meas)
                    if width > best:
                        best, best_word = width, text
                piece = []
    return best, best_word


def table_min_widths(dx: Docx, t: DTable, meas: Measurer, header=True, soft=False, orient="P"):
    """min_j колонок (§3.4 п. 1) с полями 216 и самые длинные слова. header — первая строка —
    шапка (полужирная); soft — мягкий min; картинка в ячейке — min(натуральная, 3 см), где
    натуральная ограничена высотой 0,8 текстовой области ориентации orient (§3.6)."""
    ncol = len(t.grid) or max((len(r) for r in t.rows), default=0)
    mins = [216.0] * ncol
    words = [""] * ncol
    for ri, row in enumerate(t.rows):
        col = 0
        for tc in row:
            span = to_int(wa(tc.find("w:tcPr/w:gridSpan", NS), "val")) or 1
            if span == 1 and col < ncol:
                wdt, wd = cell_min_word(dx, tc, header and ri == 0, meas, soft)
                for dr in tc.iter(q("drawing")):
                    got = picture_widths(dx, dr, orient)
                    if got and min(got[1], 1701) > wdt:
                        wdt, wd = min(got[1], 1701), "[картинка]"
                if wdt + 216 > mins[col]:
                    mins[col], words[col] = wdt + 216, wd
            col += span
    return mins, words


def table_group(dx: Docx, md_tables, t: DTable, head=True):
    """Части таблицы md, в которую входит docx-таблица t: [(docx-таблица, первая строка — шапка)].
    Без пары с md — только t, шапка — по head."""
    if md_tables is None or len(md_tables) != len(dx.tables):
        return [(t, head)]
    g = md_tables[t.index - 1].group
    return [(u, md_tables[u.index - 1].head) for u in dx.tables if md_tables[u.index - 1].group == g]


def group_min_widths(dx: Docx, group, meas: Measurer, soft=False, orient="P"):
    """min_j колонок и самые длинные слова по всем частям таблицы: сетка у частей одна (§3.4)."""
    mins, words = table_min_widths(dx, group[0][0], meas, group[0][1], soft, orient)
    for u, head in group[1:]:
        more, more_words = table_min_widths(dx, u, meas, head, soft, orient)
        for j in range(min(len(mins), len(more))):
            if more[j] > mins[j]:
                mins[j], words[j] = more[j], more_words[j]
    return mins, words


def picture_widths(dx: Docx, drawing, orient="P"):
    """(ширина рисунка в docx, натуральная ширина файла при 96 dpi, ограниченная высотой
    0,8 текстовой области ориентации orient, §3.6) в twips, или None."""
    ext = drawing.find(f".//{{{WP_NS}}}extent")
    blip = drawing.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
    rel = dx.rel_by_id.get(blip.get("{%s}embed" % R_NS)) if blip is not None else None
    data = dx.parts.get("word/" + rel.get("Target")) if rel is not None else None
    if ext is None or data is None:
        return None
    from PIL import Image
    with Image.open(io.BytesIO(data)) as im:
        w_, h_ = im.size[0] * 15, im.size[1] * 15
    cap = 0.8 * TEXT_HEIGHT.get(orient, TEXT_HEIGHT["P"])
    return (to_int(ext.get("cx")) or 0) / 635, (w_ * min(1, cap / h_) if h_ else w_)


def _is_spacer(dx: Docx, p) -> bool:
    """Отбивка 6 pt перед таблицей (§3.4): before=0 after=0 line=120 exact и keepNext."""
    sp = {a: dx.eff_ppr(p, "spacing", a) for a in ("before", "after", "line", "lineRule")}
    return (to_int(sp["before"]) or 0) == 0 and to_int(sp["after"]) == 0 and \
        to_int(sp["line"]) == 120 and sp["lineRule"] == "exact" and is_on(dx.eff_ppr(p, "keepNext"))


def _fill_level(xs, avail):
    """Порог T: Σ min(x_j, T) = avail (§3.4 п. 3); бесконечность, если Σx ≤ avail."""
    rest, k = avail, len(xs)
    for x in sorted(xs):
        if x * k > rest:
            return rest / k
        rest -= x
        k -= 1
    return float("inf")


def _align_ok(md_align: str, jc) -> bool:
    if md_align == "center":
        return jc == "center"
    if md_align == "right":
        return jc in ("right", "end")
    return jc in (None, "left", "start")


def check_tables(rep, dx: Docx, md: MdModel | None, auto_landscape=True, invariants_only=False):
    meas = measurer()
    md_tables = md.tables if md is not None else None
    pairable = md_tables is not None and len(md_tables) == len(dx.tables)
    for t in dx.tables:
        sec = dx.sections[t.section]
        lab = f"таблица {t.index} (раздел {t.section + 1}, {RU_ORIENT.get(sec.orient, '?')})"
        tbl = t.el
        tblPr = tbl.find(q("tblPr"))
        if tblPr is None:
            rep.err("table-format", f"{lab}: нет tblPr")
            continue
        bad = []
        tw_el, ind_el = tblPr.find(q("tblW")), tblPr.find(q("tblInd"))
        tblw = to_int(wa(tw_el, "w")) if wa(tw_el, "type") == "dxa" else None
        if wa(tw_el, "type") != "dxa":
            bad.append(f"tblW type={wa(tw_el, 'type')} (нужно dxa)")
        ind = 0
        if ind_el is not None:
            if wa(ind_el, "type") not in ("dxa", None):
                bad.append(f"tblInd type={wa(ind_el, 'type')} (нужно dxa)")
            ind = to_int(wa(ind_el, "w")) or 0
        W_ = sec.text_width
        md_t = md_tables[t.index - 1] if pairable else None
        if md_t is not None and not md_t.nested and ind != 0:
            bad.append(f"tblInd={ind} у таблицы верхнего уровня (нужно 0: левая граница — по "
                       f"краю текста)")
        elif ind % 360 != 0 or ind < 0:
            bad.append(f"tblInd={ind} — не отступ контейнера (кратно 360): левая граница не "
                       f"совпадает с краем текста")
        if W_ is not None and tblw is not None and tblw != W_ - ind:
            diff = tblw - (W_ - ind)
            bad.append(f"tblW={tblw} ≠ ширина текста {W_} − tblInd {ind} = {W_ - ind} "
                       f"({'шире текста на' if diff > 0 else 'уже на'} {abs(diff)})")
        if wa(tblPr.find(q("tblLayout")), "type") != "fixed":
            bad.append("нет tblLayout type=fixed")
        brd = tblPr.find(q("tblBorders"))
        for side in BORDER_SIDES:
            el = None
            for nm in SIDE_ALIASES.get(side, (side,)):
                el = brd.find(q(nm)) if brd is not None else None
                if el is not None:
                    break
            got = (wa(el, "val"), wa(el, "sz"), wa(el, "space"), (wa(el, "color") or "").upper() or None)
            if got != ("single", "4", "0", "CCCCCC"):
                bad.append(f"граница {side}={got}")
        mar = tblPr.find(q("tblCellMar"))
        for side, v in CELL_MAR.items():
            el = None
            for nm in SIDE_ALIASES.get(side, (side,)):
                el = mar.find(q(nm)) if mar is not None else None
                if el is not None:
                    break
            if to_int(wa(el, "w")) != v or wa(el, "type") not in ("dxa", None):
                bad.append(f"tblCellMar {side}={wa(el, 'w')} (нужно {v})")
        look = tblPr.find(q("tblLook"))
        got = {k: wa(look, k) for k in TBL_LOOK}
        tr0 = t.rows[0][0].getparent() if t.rows and t.rows[0] else None
        if md_t is not None:
            hl = not md_t.head  # пустая шапка или часть таблицы после разрыва
        else:  # §3.4: таблица без шапки — firstRow=0, у первой строки нет заливки шапки
            hl = got["firstRow"] == "0" and tr0 is not None \
                and not any((wa(tc.find("w:tcPr/w:shd", NS), "fill") or "").upper() == "F2F2F2"
                            for tc in t.rows[0])
        if got != dict(TBL_LOOK, firstRow="0" if hl else "1") and not \
                (hl and got == dict(TBL_LOOK, firstRow="0", val="0480")):  # val без бита firstRow
            bad.append(f"tblLook {got}" + (" (таблица без шапки: нужно firstRow=0)" if hl else ""))
        if sum(t.grid) != tblw:
            bad.append(f"сумма gridCol {sum(t.grid)} ≠ tblW {tblw}")
        if any(g <= 0 for g in t.grid) or (tblw is not None and tblw <= 0):
            bad.append(f"ширина не положительная: tblW={tblw}, gridCol {t.grid[:8]}")
        if ind > MAX_TBL_IND:
            rep.err("indent", f"{lab}: tblInd={ind} > {MAX_TBL_IND} — вложенность глубже 8 уровней "
                              f"не ограничена (§3.8)")
        ncol = len(t.grid)
        no_cant, no_tc, wrong_cnt, para_bad, jc_bad, sz_bad = 0, 0, 0, 0, 0, 0
        jc_examples, tall = [], []
        heights = []  # оценка высоты строк: строке выше страницы cantSplit не положен (§3.4)
        head = not hl  # первая строка — шапка
        for ri, row in enumerate(t.rows):
            col, h = 0, 0.0
            for tc in row:
                span = to_int(wa(tc.find("w:tcPr/w:gridSpan", NS), "val")) or 1
                h = max(h, cell_height(dx, tc, head and ri == 0, meas, sum(t.grid[col:col + span]) - 216))
                col += span
            heights.append(h)
        room = TEXT_HEIGHT[sec.orient] - (heights[0] if head else 0) \
            if heights and sec.orient in TEXT_HEIGHT else None
        cants, keeps = [], []  # cantSplit строки; keepNext у всех / у какого-нибудь абзаца строки
        for ri, row in enumerate(t.rows):
            tr = row[0].getparent() if row else None
            trpr = tr.find(q("trPr")) if tr is not None else None
            cant = trpr is not None and is_on(trpr.find(q("cantSplit")))
            kn = [is_on(dx.eff_ppr(p, "keepNext")) for tc in row for p in tc.findall(q("p"))]
            cants.append(cant)
            keeps.append((bool(kn) and all(kn), any(kn)))
            if cant and (ri or not head) and room is not None and heights[ri] > 1.1 * room:
                tall.append(ri + 1)
            elif not cant and ((head and not ri) or room is None or heights[ri] < 0.8 * room):
                no_cant += 1
            if trpr is not None and is_on(trpr.find(q("tblHeader"))):
                bad.append(f"строка {ri + 1} с tblHeader — шапка на страницах не повторяется "
                           f"(§3.4)")
            col = 0
            for ci, tc in enumerate(row):
                tcpr = tc.find(q("tcPr"))
                span = to_int(wa(tcpr.find(q("gridSpan")) if tcpr is not None else None, "val")) or 1
                tcw = tcpr.find(q("tcW")) if tcpr is not None else None
                exp_w = sum(t.grid[col:col + span])
                if wa(tcw, "type") != "dxa" or to_int(wa(tcw, "w")) != exp_w:
                    no_tc += 1
                shd = tcpr.find(q("shd")) if tcpr is not None else None
                if ri == 0 and not head and (wa(shd, "fill") or "").upper() == "F2F2F2":
                    bad.append(f"шапки нет (в md она пуста или это часть таблицы после разрыва), а ячейка "
                               f"{ci + 1} первой строки — с заливкой шапки")
                if ri == 0 and head:
                    got = (wa(shd, "val"), wa(shd, "color"), (wa(shd, "fill") or "").upper() or None)
                    if got != ("clear", "auto", "F2F2F2"):
                        bad.append(f"шапка, ячейка {ci + 1}: заливка {got}, нужно clear/auto/F2F2F2")
                    for r in tc.iter(q("r")):
                        if para_text(r).strip() and not is_on(dx.eff_rpr(r, tc.find(q('p')), "b")):
                            bad.append(f"шапка, ячейка {ci + 1}: текст «{short(para_text(r), 30)}» "
                                       f"не полужирный")
                            break
                md_align = ""
                if md_t is not None and col < len(md_t.aligns):
                    md_align = md_t.aligns[col]
                for p in tc.findall(q("p")):
                    b4 = to_int(dx.eff_ppr(p, "spacing", "before")) or 0
                    af = to_int(dx.eff_ppr(p, "spacing", "after"))
                    ln = to_int(dx.eff_ppr(p, "spacing", "line"))
                    lr = dx.eff_ppr(p, "spacing", "lineRule")
                    if b4 != 0 or af != 0 or ln != 240 or lr not in (None, "auto"):
                        para_bad += 1
                    jc = dx.eff_ppr(p, "jc", "val")
                    if jc in JUSTIFY or (md_t is not None and not _align_ok(md_align, jc)):
                        jc_bad += 1
                        if len(jc_examples) < 3:
                            jc_examples.append(f"строка {ri + 1}, колонка {col + 1}: jc={jc}, "
                                               f"в md «{md_align or 'по умолчанию'}»")
                    for r in p.iter(q("r")):
                        if not para_text(r).strip():
                            continue
                        rsz = to_int(wa(dx.eff_rpr(r, p, "sz"), "val"))
                        font = dx.eff_font(r, p, "ascii")
                        if font in CODE_FONTS:
                            if rsz not in (20, None):
                                sz_bad += 1
                        elif rsz not in (22, None):
                            sz_bad += 1
                col += span
            if col != ncol:
                wrong_cnt += 1
        nrows = len(t.rows)
        if no_cant:
            bad.append(f"нет cantSplit у {no_cant} из {nrows} строк")
        if tall:
            bad.append(f"строки {tall[:5]} выше страницы (оценка), но с cantSplit: LibreOffice "
                       f"выведет текст за нижнее поле, Word для Mac перенесёт строку целиком")
        # §3.4: keepNext — шапка с первой строкой тела (без шапки — первая строка со второй) и
        # предпоследняя строка с последней; строке без cantSplit (её можно рвать) — не ставится
        n = len(t.rows)
        need = ({0, n - 2} if n >= 2 else set()) if hl else \
            ({0} if n >= 2 else set()) | ({1, n - 2} if n >= 3 else set())
        may = need | ({n - 1} if n <= (1 if hl else 2) else set())  # строка тела одна — по вкусу

        def fits(ri):  # цепочка keepNext до строки ri + 1 влезает: от начала таблицы — в ¾ области
            s = ri     # текста, хвост — во всю область (шапка не повторяется); оценка высот своя —
            # требуем с запасом 10 %
            while s - 1 in need and cants[s - 1] and keeps[s - 1][0]:  # цепочка — по стоящим keepNext
                s -= 1
            cap = None if room is None else 0.75 * TEXT_HEIGHT[sec.orient] if s == 0 \
                else TEXT_HEIGHT[sec.orient]
            chain = sum(heights[s:ri + 1]) + (heights[ri + 1] if ri + 1 < n and cants[ri + 1] else 0)
            return cap is None or chain <= 0.9 * cap

        lacking = [ri + 1 for ri in sorted(need) if cants[ri] and not keeps[ri][0] and fits(ri)]
        extra = [ri + 1 for ri in range(n) if keeps[ri][1] and (ri not in may or not cants[ri])]
        if lacking:
            bad.append(f"нет keepNext у строк {lacking[:5]} — таблица может начаться или кончиться "
                       f"одной строкой на странице")
        if extra:
            bad.append(f"лишний keepNext у строк {extra[:5]} (нужен только "
                       f"{sorted(ri + 1 for ri in need if cants[ri])}, не у строк без cantSplit)")
        if no_tc:
            bad.append(f"tcW ≠ gridCol (или не dxa) у {no_tc} ячеек")
        if wrong_cnt:
            bad.append(f"у {wrong_cnt} строк число ячеек ≠ числу колонок сетки {ncol}")
        if para_bad:
            bad.append(f"{para_bad} абзацев в ячейках без before=0 after=0 line=240 auto")
        if jc_bad:
            bad.append(f"выравнивание не по колонке md в {jc_bad} абзацах ({'; '.join(jc_examples)})")
        if sz_bad:
            bad.append(f"размер текста не 11 pt в {sz_bad} прогонах")
        # разделитель после таблицы
        nxt = tbl.getnext()
        while nxt is not None and local(nxt) not in ("p", "tbl", "sectPr", "sdt"):
            nxt = nxt.getnext()
        if nxt is None or local(nxt) in ("tbl", "sectPr", "sdt"):
            bad.append(f"после таблицы нет абзаца-разделителя (дальше {local(nxt) if nxt is not None else 'конец'})"
                       + (" — Word склеит таблицы" if nxt is not None and local(nxt) == "tbl" else ""))
        elif nxt.find("w:pPr/w:sectPr", NS) is None and not is_on(nxt.find("w:pPr/w:pageBreakBefore", NS)):
            if has_content(nxt):  # абзац с разрывом страницы сам отделяет таблицу (носитель или блок)
                bad.append(f"после таблицы сразу текст «{short(norm_ws(para_text(nxt)), 30)}» — "
                           f"нет абзаца-разделителя")
            else:
                nn = nxt.getnext()
                while nn is not None and local(nn) not in ("p", "tbl", "sectPr", "sdt"):
                    nn = nn.getnext()
                tail = nn is not None and local(nn) == "sectPr"  # документ кончается таблицей
                vals = {a: dx.eff_ppr(nxt, "spacing", a) for a in ("before", "after", "line", "lineRule")}
                if (to_int(vals["before"]) or 0) != 0 or to_int(vals["after"]) != 0 or \
                        to_int(vals["line"]) != (20 if tail else 160) or vals["lineRule"] != "exact":
                    bad.append(f"абзац-разделитель после таблицы: spacing {vals}, нужно before=0 "
                               f"after=0 line={20 if tail else 160} exact"
                               + (" (хвост документа — носитель 1 pt, иначе пустая страница)" if tail else ""))
                if tail and to_int(wa(nxt.find("w:pPr/w:rPr/w:sz", NS), "val")) != 2:
                    bad.append("хвост документа после таблицы: знак абзаца не sz=2 (носитель 1 pt)")
                if nn is not None and local(nn) == "p" and nn.find("w:pPr/w:sectPr", NS) is not None \
                        and not has_content(nn):
                    bad.append("после таблицы два разделителя: пустой абзац и абзац-носитель "
                               "разрыва (второй не нужен)")
                if nn is not None and local(nn) == "p" and is_on(nn.find("w:pPr/w:pageBreakBefore", NS)):
                    bad.append("после таблицы разделитель, а за ним разрыв страницы: у таблицы, дошедшей "
                               "до низа страницы, разделитель уйдёт на пустую страницу")
        # §3.4: таблица верхнего уровня сразу за заголовком — перед ней отбивка 6 pt
        if (md_t is not None and not md_t.nested) or (md_t is None and ind == 0):
            prv, empties = tbl.getprevious(), []
            while prv is not None and local(prv) == "p" and not has_content(prv) \
                    and not para_atoms(prv) and prv.find("w:pPr/w:sectPr", NS) is None \
                    and prv.find("w:pPr/w:numPr", NS) is None and prv.find("w:pPr/w:shd", NS) is None:
                empties.append(prv)
                prv = prv.getprevious()
            if prv is not None and local(prv) == "p" and dx.heading_level(prv) is not None \
                    and has_content(prv) and not any(_is_spacer(dx, e) for e in empties):
                bad.append("стоит сразу за заголовком, но перед ней нет отбивки 6 pt (пустой абзац "
                           "before=0 after=0 line=120 exact с keepNext)")
        # части таблицы, разрезанной \newpage (§3.4): сетка одна — у частей те же gridCol и tblW
        group = table_group(dx, md_tables, t, head)
        first = group[0][0]
        if first is not t:
            fw = to_int(wa(first.el.find("w:tblPr/w:tblW", NS), "w"))
            if first.grid != t.grid or fw != tblw:
                bad.append(f"часть {md_t.part + 1} таблицы md: сетка {t.grid[:6]} (tblW {tblw}) ≠ "
                           f"сетке первой части {first.grid[:6]} (tblW {fw}) — сетка у частей одна")
        # ширины колонок против слов (§3.4): нижняя граница — мягкий min (даты и коды целиком), если
        # мягкие помещаются, иначе жёсткий; при Σmin > ширины узкие колонки (min ≤ порога T)
        # сохраняют слово целиком, дефицит — с широких. Замер конвертера — с запасом 3 %; у частей
        # таблицы — по всем частям
        mins, words = group_min_widths(dx, group, meas, orient=sec.orient)
        softs, swords = group_min_widths(dx, group, meas, soft=True, orient=sec.orient)
        if W_ is not None and len(mins) == len(t.grid) and t.grid:
            avail = W_ - ind
            sum_min = sum(mins)

            def margin(xs):
                return [(x - 216) * 1.03 + 216 for x in xs]
            soft = [max(a, b) for a, b in zip(mins, softs)]
            if sum(margin(soft)) <= avail - 2:
                need, nw = soft, [s if b > a else w for a, b, w, s in zip(mins, softs, words, swords)]
            else:
                need, nw = mins, words
            level = _fill_level(margin(need), avail)
            for j, (g, mn, mm) in enumerate(zip(t.grid, need, margin(need))):
                if mm <= level and g + 2 < mn:
                    bad.append(f"колонка {j + 1}: gridCol={g} уже слова «{nw[j]}» ({round(mn)} с "
                               f"полями) — слово разорвётся"
                               + ("" if level == float("inf") else
                                  f" (таблица шире страницы, но колонка узкая: порог {round(level)})"))
            if sum(margin(mins)) <= avail - 2:
                # картинка в ячейке: колонка даёт ей натуральную ширину или хотя бы 3 см (§3.4)
                for row in t.rows:
                    col = 0
                    for tc in row:
                        for dr in tc.iter(q("drawing")):
                            got = picture_widths(dx, dr, sec.orient)
                            if got and got[0] + 15 < min(got[1], 1701):
                                bad.append(f"картинка в колонке {col + 1}: {got[0] / 567:.2f} см при "
                                           f"натуральной {got[1] / 567:.2f} см — колонка её не учла")
                        col += to_int(wa(tc.find("w:tcPr/w:gridSpan", NS), "val")) or 1
            # §4.3: широкая таблица на книжной странице (замер жёсткий, картинки — как на книжной)
            mins_p = mins if sec.orient == "P" else group_min_widths(dx, group, meas)[0]
            if sec.orient == "P" and auto_landscape and sum_min > TEXT_WIDTH["P"] - ind:
                msg = (f"{lab}: Σmin={round(sum_min)} > {TEXT_WIDTH['P'] - ind} — на книжной "
                       f"странице не помещается без разрыва слов, но осталась книжной")
                if invariants_only or md_t is None:
                    rep.warn("auto-landscape", msg)
                else:
                    rep.err("auto-landscape", msg)
            if md_t is not None and not md_t.nested and md_t.region == "P" and auto_landscape \
                    and sec.orient == "L" and (sum(m * 1.03 for m in mins_p)) <= TEXT_WIDTH["P"] - ind - 50:
                rep.err("auto-landscape", f"{lab}: помещается на книжной странице "
                                          f"(Σmin≈{round(sum(mins_p))}), но ушла в альбомный раздел")
            if md_t is not None and md_t.region == "L" and sec.orient != "L":
                rep.err("auto-landscape", f"{lab}: участок размечен альбомным, а таблица в "
                                          f"книжном разделе")
        if bad:
            for b in bad:
                rep.err("table-format", f"{lab}: {b}")


def check_table_content(rep, dx: Docx, md: MdModel):
    if len(md.tables) != len(dx.tables):
        rep.err("tables", f"таблиц в docx {len(dx.tables)}, в md {len(md.tables)}")
    for k, (mt, dt) in enumerate(zip(md.tables, dx.tables), 1):
        lab = f"таблица {k}"
        if mt.headerless:  # §3.4: пустая шапка в docx не выводится
            mt = MdTable(mt.aligns, mt.rows[1:])
        mrows, ncols = len(mt.rows), len(mt.aligns)
        drows = len(dt.rows)
        dcols = max((len(r) for r in dt.rows), default=0)
        if (mrows, ncols) != (drows, dcols):
            rep.err("tables", f"{lab}: размер {drows}×{dcols}, в md {mrows}×{ncols}")
        diffs = []
        for ri, (mr, dr) in enumerate(zip(mt.rows, dt.rows)):
            for ci, (mc, dc) in enumerate(zip(mr, dr)):
                a = norm_lines(mc)
                b = norm_lines(PLACEHOLDER_RE.sub(" ", cell_text(dc)))
                if a != b:
                    diffs.append(f"[{ri + 1},{ci + 1}] docx «{short(' ⏎ '.join(b), 50)}» ≠ md "
                                 f"«{short(' ⏎ '.join(a), 50)}»")
        if diffs:
            more = f" … и ещё {len(diffs) - 5}" if len(diffs) > 5 else ""
            rep.err("tables", f"{lab}: тексты ячеек расходятся: " + "; ".join(diffs[:5]) + more)


# --------------------------------------------------------------------------
# §4 план разделов по md: маркеры, флаг, автоальбомные группы
# --------------------------------------------------------------------------


def _lead_heading(vis, s):
    """§4.3 п. 2 и 4: назад от блока s — не больше 2 коротких простых абзацев и 1 линии, затем
    заголовок → его номер; иначе (или разрыв, начало участка) — None."""
    paras = hrs = 0
    j = s
    while j > 0 and not vis[j].barrier:
        j -= 1
        b = vis[j]
        if b.kind == "heading":
            return j
        if b.kind == "hr" and hrs < 1:
            hrs += 1
        elif b.simple and paras < 2:
            paras += 1
        else:
            return None
    return None


def _pull_start(vis, w):
    """§4.3: первый блок автоальбомной группы с широким блоком w — подпись с двоеточием, подводка
    с заголовком, заголовки подряд, подъём к строго старшему заголовку."""
    s = w
    if s > 0 and not vis[s].barrier and vis[s - 1].kind == "paragraph" and vis[s - 1].colon:
        s -= 1  # п. 1: подпись — всегда
    h, top = _lead_heading(vis, s), 99
    while h is not None:
        s, top = h, min(top, vis[h].level)
        while s > 0 and not vis[s].barrier and vis[s - 1].kind == "heading":  # п. 3
            s -= 1
            top = min(top, vis[s].level)
        h = _lead_heading(vis, s)
        if h is not None and vis[h].level >= top:  # п. 4: только строго старший
            h = None
    return s


def docx_flow(dx: Docx, n_images: int):
    """Слова docx по порядку; номера слов, с которых начинаются разделы и стоят разрывы страниц."""
    out = []
    for s in dx.sections:
        out.append(" \x01 ")
        for el in s.elements:
            for p in [el] if el.tag == q("p") else el.iter(q("p")):
                if is_on(p.find("w:pPr/w:pageBreakBefore", NS)):
                    out.append(" \x00 ")
                for x in p.iter(q("t"), q("tab"), q("br"), q("cr"), q("noBreakHyphen")):
                    if x.tag == q("t"):
                        out.append(x.text or "")
                    elif x.tag == q("br") and wa(x, "type") == "page":
                        out.append(" \x00 ")
                    elif x.tag == q("noBreakHyphen"):
                        out.append("-")
                    elif x.getparent().tag != q("tabs"):
                        out.append(" ")
                out.append("\n")
    words, secs, brks = [], [], []
    for w in PLACEHOLDER_RE.sub(" ", "".join(out), count=n_images).split() if n_images \
            else "".join(out).split():
        if w == "\x01":
            secs.append(len(words))
        elif w == "\x00":
            brks.append(len(words))
        else:
            words.append(w)
    return words, secs, brks


def check_plan(rep, dx: Docx, md: MdModel, auto_landscape=True):
    """§4: ожидаемые разделы и разрывы по md — маркеры (§4.1), флаг (§4.2), автоальбомные группы
    с подтягиванием и склейкой (§4.3) — против разделов и разрывов docx."""
    vis = [b for b in md.flow if isinstance(b, MdBlock) and b.visible]
    if not vis or len(md.tables) != len(dx.tables):
        return  # без пары таблиц широкие не определить; расхождение уже в «tables»
    meas, wide = measurer(), set()
    for t in dx.tables:
        ind = to_int(wa(t.el.find("w:tblPr/w:tblInd", NS), "w")) or 0
        mins, _ = group_min_widths(dx, table_group(dx, md.tables, t), meas)  # по всем частям таблицы
        if sum(mins) > TEXT_WIDTH["P"] - ind:
            wide.add(t.index - 1)  # не влезает и по замеру без запаса — широкая наверняка
        elif sum(x * 1.03 for x in mins) > TEXT_WIDTH["P"] - ind - 50 \
                and dx.sections[t.section].orient == "L":
            wide.add(t.index - 1)  # пограничная: решение конвертера принимается
    in_group, prev = [False] * len(vis), None
    for w, b in enumerate(vis):
        if not (auto_landscape and b.region == "P" and wide.intersection(b.tables)):
            continue
        s = _pull_start(vis, w)
        if prev is not None and len(vis[prev + 1:s]) <= 2 and not any(x.tables for x in vis[prev + 1:s]) \
                and not any(x.barrier for x in vis[prev + 1:s + 1]):
            s = prev + 1  # склейка: между группами не больше 2 видимых блоков, нет таблиц и разрыва
        for k in range(s, w + 1):
            in_group[k] = True
        prev = w
    exp_sec, exp_brk, pos, cur = [], [], 0, None
    for k, b in enumerate(vis):
        eff = "L" if in_group[k] else b.region
        if eff != cur:
            exp_sec.append((eff, pos))
            cur = eff
        elif b.barrier:
            exp_brk.append(pos)  # смена участка без смены ориентации — тоже новая страница
        exp_brk += [pos + w for w in b.breaks]  # части таблицы, разрезанной \newpage (§3.4)
        pos += b.words
    words, secs, brks = docx_flow(dx, len(md.images))
    ru = {"P": "К", "L": "А"}
    got = [s.orient for s in dx.sections]
    # §4.1: раздел той же ориентации, что и предыдущий, — та же новая страница, что и разрыв
    same = {k for k in range(1, len(got)) if got[k] == got[k - 1]}
    brks += [secs[k] for k in sorted(same)]
    got = [o for k, o in enumerate(got) if k not in same]
    secs = [p for k, p in enumerate(secs) if k not in same]
    if [o for o, _ in exp_sec] != got:
        rep.err("plan", f"разделы по md {', '.join(ru[o] for o, _ in exp_sec)} — в docx "
                        f"{', '.join(ru.get(o, '?') for o in got)}")
    if words != md.text.split():
        return  # слова расходятся — места разрывов не сопоставить (расхождение в «words»)

    def at(n):
        return f"«{short(' '.join(words[n:n + 6]), 40)}»" if n < len(words) else "конец документа"
    if [o for o, _ in exp_sec] == got:
        for k, ((o, p), d) in enumerate(zip(exp_sec, secs)):
            if p != d:
                rep.err("plan", f"раздел {k + 1} ({ru[o]}) по md начинается со слов {at(p)}, "
                                f"в docx — {at(d)}")
    exp, have = Counter(exp_brk), Counter(brks)
    for p in sorted((exp - have).elements()):
        rep.err("plan", f"нет разрыва страницы перед {at(p)}")
    for p in sorted((have - exp).elements()):
        rep.err("plan", f"лишний разрыв страницы перед {at(p)}")


# --------------------------------------------------------------------------
# Нумерация: §5 (numbering.xml) и §3.5
# --------------------------------------------------------------------------


@dataclass
class NumInfo:
    num_ids: dict  # numId → (abstractNumId, {ilvl: startOverride})
    abstracts: dict  # abstractNumId → {ilvl: lvl-элемент}


def numbering_info(dx: Docx) -> NumInfo:
    info = NumInfo({}, {})
    if dx.numbering is None:
        return info
    for an in dx.numbering.findall(q("abstractNum")):
        lv = {}
        for lvl in an.findall(q("lvl")):
            lv[to_int(wa(lvl, "ilvl"))] = lvl
        info.abstracts[wa(an, "abstractNumId")] = lv
    for num in dx.numbering.findall(q("num")):
        so = {}
        for ov in num.findall(q("lvlOverride")):
            s = ov.find(q("startOverride"))
            if s is not None:
                so[to_int(wa(ov, "ilvl"))] = to_int(wa(s, "val"))
        info.num_ids[wa(num, "numId")] = (wa(num.find(q("abstractNumId")), "val"), so)
    return info


def numbered_paragraphs(dx: Docx):
    out = []
    for p in dx.body.iter(q("p")):
        np_ = p.find("w:pPr/w:numPr", NS)
        if np_ is None:
            continue
        nid = wa(np_.find(q("numId")), "val")
        if nid in (None, "0"):
            continue
        out.append((p, nid, to_int(wa(np_.find(q("ilvl")), "val")) or 0))
    return out


def check_numbering(rep, dx: Docx):
    paras = numbered_paragraphs(dx)
    if dx.numbering is None:
        if paras:
            rep.err("numbering", "абзацы ссылаются на нумерацию, а numbering.xml нет")
        return
    ids = [wa(n, "numId") for n in dx.numbering.findall(q("num"))]
    aids = [wa(a, "abstractNumId") for a in dx.numbering.findall(q("abstractNum"))]
    for name, lst in (("numId", ids), ("abstractNumId", aids)):
        dup = [k for k, v in Counter(lst).items() if v > 1]
        if dup:
            rep.err("numbering", f"в numbering.xml повторяются {name}: {dup}")
    info = numbering_info(dx)
    for nid, (aid, _) in info.num_ids.items():
        if aid not in info.abstracts:
            rep.err("numbering", f"num {nid} ссылается на несуществующий abstractNum {aid}")
    used_abs = set()
    for p, nid, ilvl in paras:
        if nid not in info.num_ids:
            rep.err("numbering", f"абзац «{short(para_text(p), 30)}» ссылается на несуществующий "
                                 f"numId {nid}")
            continue
        used_abs.add((info.num_ids[nid][0], ilvl))
    bad = set()
    for aid, ilvl in sorted(used_abs, key=lambda x: (str(x[0]), x[1])):
        lvl = info.abstracts.get(aid, {}).get(ilvl)
        if lvl is None:
            rep.err("numbering", f"abstractNum {aid}: нет уровня {ilvl}")
            continue
        fmt = wa(lvl.find(q("numFmt")), "val")
        txt = wa(lvl.find(q("lvlText")), "val")
        ind = lvl.find("w:pPr/w:ind", NS)
        left = to_int(wa(ind, "left") if wa(ind, "left") is not None else wa(ind, "start"))
        hanging = to_int(wa(ind, "hanging"))
        if left != 360 * (ilvl + 1) or hanging != 360:
            bad.add(f"abstractNum {aid} уровень {ilvl}: ind left={left} hanging={hanging}, нужно "
                    f"left={360 * (ilvl + 1)} hanging=360")
        if fmt == "bullet":
            if txt != BULLETS[ilvl % 3]:
                bad.add(f"abstractNum {aid} уровень {ilvl}: маркер «{txt}», нужно «{BULLETS[ilvl % 3]}»")
        elif fmt == "decimal":
            if txt not in (f"%{ilvl + 1}.", f"%{ilvl + 1})"):
                bad.add(f"abstractNum {aid} уровень {ilvl}: lvlText «{txt}», нужно «%{ilvl + 1}.» "
                        f"или «%{ilvl + 1})»")
        else:
            bad.add(f"abstractNum {aid} уровень {ilvl}: numFmt={fmt}, нужно bullet или decimal")
    for b in sorted(bad):
        rep.err("numbering", b)


def check_lists(rep, dx: Docx, md: MdModel):
    paras = numbered_paragraphs(dx)
    if len(paras) != len(md.items):
        rep.err("lists", f"пунктов списков с нумерацией в docx {len(paras)}, в md {len(md.items)}")
        return
    info = numbering_info(dx)
    lists = {ml.id: ml for ml in md.lists}
    list_to_num, num_to_list = {}, {}
    problems = []
    for (p, nid, ilvl), item in zip(paras, md.items):
        lst = lists[item.list_id]
        lab = f"пункт «{short(norm_ws(item.text), 30)}»"
        if norm_ws(PLACEHOLDER_RE.sub(" ", para_text(p))) != norm_ws(item.text):
            problems.append(f"{lab}: в docx на этом месте «{short(norm_ws(para_text(p)), 30)}»")
            continue
        if ilvl != min(item.depth, 8):  # уровней нумерации в Word 9 (0…8), глубже — ilvl 8
            problems.append(f"{lab}: ilvl={ilvl}, глубина в md {item.depth}")
        if list_to_num.setdefault(lst.id, nid) != nid:
            problems.append(f"{lab}: один список md, а numId разные ({list_to_num[lst.id]} и {nid})")
        if num_to_list.setdefault(nid, lst.id) != lst.id:
            problems.append(f"{lab}: numId {nid} уже занят другим списком — нумерация продолжится, "
                            f"а не начнётся заново")
        if not lst.in_quote:
            sid = wa(p.find("w:pPr/w:pStyle", NS), "val")
            if sid not in (None, "Normal", dx.default_pstyle):
                problems.append(f"{lab}: стиль {sid}, нужен Normal + numPr")
        aid, so = info.num_ids.get(nid, (None, {}))
        lvl = info.abstracts.get(aid, {}).get(ilvl)
        fmt = wa(lvl.find(q("numFmt")), "val") if lvl is not None else None
        if lst.ordered:
            if fmt != "decimal":
                problems.append(f"{lab}: нумерованный в md, а numFmt={fmt}")
            want = f"%{ilvl + 1}{lst.delim}"
            txt = wa(lvl.find(q("lvlText")), "val") if lvl is not None else None
            if txt != want:
                problems.append(f"{lab}: lvlText «{txt}», нужно «{want}»")
            if so.get(ilvl) != lst.start:
                problems.append(f"{lab}: startOverride уровня {ilvl} = {so.get(ilvl)}, нужно "
                                f"{lst.start} (номер начала списка в md)")
        elif fmt != "bullet":
            problems.append(f"{lab}: маркированный в md, а numFmt={fmt}")
    uniq = list(dict.fromkeys(problems))
    for pr in uniq[:12]:
        rep.err("lists", pr)
    if len(uniq) > 12:
        rep.err("lists", f"… и ещё {len(uniq) - 12} расхождений по спискам")


# --------------------------------------------------------------------------
# Сверка текста: слова, заголовки, разметка, картинки
# --------------------------------------------------------------------------


def docx_text(dx: Docx) -> str:
    return "\n".join(para_text(p) for p in dx.body.iter(q("p")))


LEFTOVERS = [
    ("**", r"\*\*"), ("__", r"__"), ("`", r"`"), ("~~", r"~~"),
    ("|---", r"\|\s*:?-{3,}"), ("](", r"\]\("), ("<!--", r"<!--"), ("-->", r"-->"),
    ("```", r"```"), ("~~~", r"~~~"), ("<br>", r"<br\s*/?>"),
    ("HTML-сущность", r"&(?:[A-Za-z][A-Za-z0-9]*|#\d+|#[xX][0-9A-Fa-f]+);"),
    ("HTML-тег", r"</?[A-Za-z][A-Za-z0-9-]*(?:\s[^<>]*)?/?>"),
    ("# в начале строки", r"(?m)^[ \t]*#{1,6}(?:[ \t]|$)"),
    ("> в начале строки", r"(?m)^[ \t]*>"),
    ("маркер списка в начале строки", r"(?m)^[ \t]*[-*+][ \t]"),
    ("номер списка в начале строки", r"(?m)^[ \t]*\d{1,9}[.)][ \t]"),
    ("\\newpage/\\pagebreak", r"\\(?:newpage|pagebreak)\b"),
]


def check_text(rep, dx: Docx, md: MdModel):
    dtext = docx_text(dx)
    n_ph = len(PLACEHOLDER_RE.findall(dtext))
    dtext_words = PLACEHOLDER_RE.sub(" ", dtext, count=len(md.images)) if md.images else dtext
    mw, dw = Counter(md.text.split()), Counter(dtext_words.split())
    rep.info.append(f"Слов: docx {sum(dw.values())}, md {sum(mw.values())}")
    if mw != dw:
        miss, extra = mw - dw, dw - mw
        parts = []
        if miss:
            parts.append("нет в docx: " + ", ".join(f"«{k}»×{v}" if v > 1 else f"«{k}»"
                                                   for k, v in miss.most_common(12)) +
                         (f" … всего {sum(miss.values())}" if len(miss) > 12 else ""))
        if extra:
            parts.append("лишние в docx: " + ", ".join(f"«{k}»×{v}" if v > 1 else f"«{k}»"
                                                      for k, v in extra.most_common(12)) +
                         (f" … всего {sum(extra.values())}" if len(extra) > 12 else ""))
        rep.err("words", "; ".join(parts))
    else:  # набор тот же — сверить и порядок (§5: текст не переставляется)
        a, b = md.text.split(), dtext_words.split()
        k = next((i for i, (x, y) in enumerate(zip(a, b)) if x != y), None)
        if k is not None:
            rep.err("words", f"порядок слов расходится с md, с {k + 1}-го слова: docx "
                             f"«{short(' '.join(b[k:k + 6]), 50)}», md «{short(' '.join(a[k:k + 6]), 50)}»")
    # заголовки
    dh = []
    for p in dx.body.iter(q("p")):
        lvl = dx.heading_level(p)
        if lvl is not None:
            dh.append((lvl, norm_lines(PLACEHOLDER_RE.sub(" ", para_text(p)))))
    rep.info.append(f"Заголовков: docx {len(dh)}, md {len(md.headings)}")
    if dh != md.headings:
        msgs = []
        for k in range(max(len(dh), len(md.headings))):
            a = md.headings[k] if k < len(md.headings) else None
            b = dh[k] if k < len(dh) else None
            if a != b:
                fa = f"H{a[0]} «{short(' ⏎ '.join(a[1]), 50)}»" if a else "—"
                fb = f"H{b[0]} «{short(' ⏎ '.join(b[1]), 50)}»" if b else "—"
                msgs.append(f"№{k + 1}: docx {fb}, md {fa}")
            if len(msgs) >= 4:
                break
        rep.err("headings", f"последовательность заголовков расходится (docx {len(dh)}, md "
                            f"{len(md.headings)}): " + "; ".join(msgs))
    # остатки разметки (сравнение с проекцией md: экранированные символы законны)
    for name, rx in LEFTOVERS:
        nm, nd = len(re.findall(rx, md.text)), len(re.findall(rx, dtext))
        if nd > nm:
            ex = re.search(rx, dtext)
            ctx = dtext[max(0, ex.start() - 25): ex.end() + 25] if ex else ""
            rep.err("markup", f"«{name}»: в docx {nd}, в тексте md {nm} — например «{short(ctx, 60)}»")
    # картинки
    drawings = sum(1 for _ in dx.body.iter(q("drawing")))
    if drawings + n_ph != len(md.images):
        rep.err("images", f"картинок в md {len(md.images)}, в docx рисунков {drawings} и "
                          f"заглушек {n_ph}")
    from urllib.parse import unquote, urlparse
    missing = 0
    for alt, src in md.images:  # markdown-it кодирует кириллицу и пробелы в src (%D1%81…)
        # §3.6: файл — путь от каталога md, абсолютный или `file:`; прочие схемы (http, data…) — нет
        mm = re.match(r"([a-z][a-z0-9+.-]+):", src, re.I)
        path = (urlparse(src).path if mm.group(1).lower() == "file" else None) if mm else src
        try:
            missing += path is None or not (md.path.parent / unquote(path)).is_file()
        except OSError:  # имя длиннее допустимого и т. п. — файла нет
            missing += 1
    if n_ph < missing:
        rep.err("images", f"картинок без файла (или удалённых) {missing}, заглушек «[изображение: "
                          f"…]» {n_ph}")
    # заглушка по порядку картинок md несёт alt как в md: сущности и \-экранирование (§3.6)
    seq = []
    for p in dx.body.iter(q("p")):
        s = "".join("\x02" if x.tag == q("drawing") else (x.text or "")
                    for x in p.iter(q("t"), q("drawing")))
        seq += [s[mm.start():] for mm in re.finditer(r"\x02|\[изображение", s)]
    if len(seq) == len(md.images):
        for k, (got, (alt, _)) in enumerate(zip(seq, md.images), 1):
            want = f"[изображение: {norm_ws(alt)}]"
            if got[0] != "\x02" and norm_ws(alt) and not norm_ws(got).startswith(want):
                rep.err("images", f"картинка {k}: в docx «{short(norm_ws(got), 50)}», нужно "
                                  f"«{short(want, 50)}» (alt из md)")


LINK_SCHEMES = ("http", "https", "mailto", "tel")
SCHEME_RE = re.compile(r"([a-z][a-z0-9+.-]*):", re.I)


def link_ok(href: str) -> bool:
    """§3.6: гиперссылкой становится только адрес со схемой http, https, mailto или tel."""
    mm = SCHEME_RE.match(href or "")
    return bool(mm) and mm.group(1).lower() in LINK_SCHEMES


def check_links(rep, dx: Docx, md: MdModel | None):
    """§3.6: гиперссылки — внешние, только разрешённых схем; в полной сверке — те же адреса и в том
    же порядке, что у ссылок md с этими схемами (остальные ссылки — текст без гиперссылки)."""
    from urllib.parse import unquote
    got = []
    for h in dx.body.iter(q("hyperlink")):
        lab = f"гиперссылка «{short(norm_ws(para_text(h)), 30)}»"
        rel = dx.rel_by_id.get(h.get(f"{{{R_NS}}}id") or "")
        if rel is None:
            rep.err("links", f"{lab} без внешнего адреса (якорь {h.get(q('anchor'))!r}) — якоря "
                             f"остаются текстом")
            continue
        target = rel.get("Target") or ""
        if not link_ok(target):
            rep.err("links", f"{lab} на «{short(target, 50)}» — схема не из http, https, mailto, tel")
        got.append(target)
    if md is None:
        return
    want = [x for x in md.links if link_ok(x)]
    k = next((i for i, (a, b) in enumerate(zip(got, want)) if unquote(a) != unquote(b)),
             None if len(got) == len(want) else min(len(got), len(want)))
    if k is not None:
        rep.err("links", f"гиперссылок в docx {len(got)}, ссылок md с адресом http/https/mailto/tel "
                         f"{len(want)}; №{k + 1}: docx «{short(got[k] if k < len(got) else '—', 50)}», "
                         f"md «{short(want[k] if k < len(want) else '—', 50)}»")


_GLYPHS = {}


def _glyph_fonts():
    """(Helvetica Neue, Arial) Pillow для проверки глифов или None — нет Pillow или шрифтов."""
    if "fonts" not in _GLYPHS:
        _GLYPHS["fonts"] = None
        try:
            from PIL import ImageFont
            arial = next((p for p in _ARIAL[(False, False)] if Path(p).exists()), None)
            if Path(_HN_FILE).exists() and arial:
                _GLYPHS["fonts"] = (ImageFont.truetype(_HN_FILE, 64, index=0),
                                    ImageFont.truetype(arial, 64))
        except Exception:  # noqa: BLE001
            pass
    return _GLYPHS["fonts"]


def has_glyph(font, ch: str) -> bool:
    """Глиф есть, если отрисовка символа отличается от отрисовки заведомо отсутствующего U+10FFFD."""
    from PIL import Image, ImageDraw

    def sig(c):
        im = Image.new("L", (128, 128))
        ImageDraw.Draw(im).text((32, 32), c, font=font, fill=255)
        return im.tobytes()
    key = (id(font), ch)
    if key not in _GLYPHS:
        if (id(font), None) not in _GLYPHS:
            _GLYPHS[(id(font), None)] = sig("\U0010fffd")
        _GLYPHS[key] = sig(ch) != _GLYPHS[(id(font), None)]
    return _GLYPHS[key]


def check_glyphs(rep, dx: Docx):
    """§3.2: символ, которого нет в Helvetica Neue, но есть в Arial, — в прогоне с rFonts Arial
    (ascii и hAnsi), иначе Word для Mac подставит Times; в прогоне Arial — только такие символы
    и пробелы. Код (Menlo) не трогается."""
    fonts = _glyph_fonts()
    if fonts is None:
        rep.info.append("Глифы: не проверялись — нет Pillow или шрифтов Helvetica Neue и Arial")
        return
    hn, arial = fonts
    need = {}

    def arial_only(ch):
        if ch not in need:
            need[ch] = ord(ch) > 127 and unicodedata.category(ch)[0] not in "CZ" \
                and not has_glyph(hn, ch) and has_glyph(arial, ch)
        return need[ch]
    miss, extra = Counter(), Counter()
    for p in dx.body.iter(q("p")):
        for r in p.iter(q("r")):
            txt = "".join(x.text or "" for x in r.iter(q("t")))
            if not txt.strip() or dx.eff_font(r, p, "ascii") in CODE_FONTS:
                continue
            is_arial = dx.eff_font(r, p, "ascii") == "Arial" and dx.eff_font(r, p, "hAnsi") == "Arial"
            for ch in txt:
                if arial_only(ch) and not is_arial:
                    miss[ch] += 1
                elif is_arial and not ch.isspace() and not arial_only(ch):
                    extra[ch] += 1
    if miss:
        rep.err("glyphs", "символы без глифа в Helvetica Neue вне прогона Arial (Word для Mac "
                          "нарисует их Times): " + ", ".join(f"«{c}» U+{ord(c):04X} ×{n}"
                                                             for c, n in miss.most_common(8)))
    if extra:
        rep.err("glyphs", "в прогоне Arial символы, которые есть в Helvetica Neue: "
                + ", ".join(f"«{c}»×{n}" for c, n in extra.most_common(8)))


def check_indent(rep, dx: Docx):
    """§3.8: отступ вложенных абзацев ограничен 8 уровнями — текст не уходит за правое поле."""
    deep = []
    for p in dx.body.iter(q("p")):
        ind = p.find("w:pPr/w:ind", NS)
        v = max(to_int(wa(ind, "left")) or 0, to_int(wa(ind, "start")) or 0)
        if v > MAX_PARA_IND:
            deep.append((v, short(norm_ws(para_text(p)), 30)))
    if deep:
        rep.err("indent", f"{len(deep)} абзацев с отступом больше {MAX_PARA_IND} (глубже 8 уровней), "
                          f"наибольший {max(deep)[0]} — «{max(deep)[1]}»")


def check_images(rep, dx: Docx):
    """§3.6: картинка не шире текста раздела и не выше 0,8 его текстовой области."""
    for k, dr in enumerate(dx.body.iter(q("drawing")), 1):
        ext = dr.find(f".//{{{WP_NS}}}extent")
        o = dx.sections[dx.section_index(dr)].orient
        if ext is None or o not in TEXT_WIDTH:
            continue
        cx, cy = (to_int(ext.get("cx")) or 0) / 635, (to_int(ext.get("cy")) or 0) / 635
        if cx > TEXT_WIDTH[o] + 15 or cy > 0.8 * TEXT_HEIGHT[o] + 15:
            rep.err("images", f"рисунок {k}: {cx / 567:.1f}×{cy / 567:.1f} см — больше ширины текста "
                              f"{TEXT_WIDTH[o] / 567:.1f} см или 0,8 высоты {0.8 * TEXT_HEIGHT[o] / 567:.1f} см "
                              f"({RU_ORIENT.get(o, o)} раздел)")


# --------------------------------------------------------------------------
# Сборка
# --------------------------------------------------------------------------


def run_checks(md_path=None, docx_path=None, expect_orient=None, landscape=False,
               auto_landscape=True, invariants_only=False) -> Report:
    rep = Report(md_path=str(md_path) if (md_path and not invariants_only) else None,
                 docx_path=str(docx_path))
    try:
        dx = Docx(docx_path)
    except Exception as e:  # noqa: BLE001
        rep.err("structure", f"docx не читается: {e}")
        return rep
    md = None
    if not invariants_only:
        try:
            md = parse_md(md_path, landscape=landscape)
        except MdInputError as e:
            rep.err("input", f"md: {e} — конвертер обязан был отказать (§2), а docx есть")
            return rep
        except Exception as e:  # noqa: BLE001
            rep.err("internal", f"md не читается: {e}")
            return rep
        if md.too_deep:
            rep.warn("input", "md глубже 100 уровней вложенности: хвост не разобран (maxNesting) — "
                              "конвертер должен предупредить, сверка — по разобранной части")
    expect = parse_orient_list(expect_orient) if isinstance(expect_orient, str) else expect_orient
    steps = [
        ("структура", lambda: check_structure(rep, dx)),
        ("пакет", lambda: check_package(rep, dx)),
        ("порядок §5", lambda: check_order(rep, dx)),
        ("разделы", lambda: check_sections(rep, dx, expect, md is not None and not md.text.strip())),
        ("docDefaults", lambda: check_defaults(rep, dx)),
        ("fontTable", lambda: check_fonts(rep, dx)),
        ("свойства", lambda: check_core(rep, dx, None if md is None else
                                       (md.title or Path(md_path).stem))),  # пустой H1 — как нет H1
        ("стили заголовков", lambda: check_heading_styles(rep, dx)),
        ("абзацы заголовков", lambda: check_heading_paragraphs(rep, dx)),
        ("таблицы", lambda: check_tables(rep, dx, md, auto_landscape, invariants_only)),
        ("нумерация", lambda: check_numbering(rep, dx)),
        ("ссылки", lambda: check_links(rep, dx, md)),
        ("глифы", lambda: check_glyphs(rep, dx)),
        ("отступы", lambda: check_indent(rep, dx)),
        ("картинки", lambda: check_images(rep, dx)),
    ]
    if md is not None:
        steps += [
            ("текст", lambda: check_text(rep, dx, md)),
            ("таблицы md", lambda: check_table_content(rep, dx, md)),
            ("списки md", lambda: check_lists(rep, dx, md)),
            ("план разделов", lambda: check_plan(rep, dx, md, auto_landscape)),
        ]
    for name, fn in steps:
        try:
            fn()
        except Exception:  # noqa: BLE001
            tb = traceback.format_exc(limit=3).strip().splitlines()
            rep.err("internal", f"проверка «{name}» упала: {tb[-1]}")
    rep.info.insert(0, f"Таблиц: docx {len(dx.tables)}" + (f", md {len(md.tables)}" if md else "")
                    + f"; замер ширины: {measurer().source}")
    return rep


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="verify.py",
        description="Независимый сверщик md ↔ docx для md2docx (SPEC.md §6).")
    ap.add_argument("files", nargs="+", metavar="FILE",
                    help="IN.md OUT.docx; с --invariants-only достаточно OUT.docx")
    ap.add_argument("--expect-orient", metavar="P,L,P",
                    help="ожидаемая ориентация разделов по порядку (P/L или К/А)")
    ap.add_argument("--invariants-only", action="store_true",
                    help="только инварианты docx (§3–§5), без сверки текста с md")
    ap.add_argument("--landscape", action="store_true",
                    help="документ собран с --landscape (ориентация по умолчанию альбомная)")
    ap.add_argument("--no-auto-landscape", action="store_true",
                    help="документ собран с --no-auto-landscape")
    args = ap.parse_args(argv)
    if args.invariants_only:
        if len(args.files) not in (1, 2):
            ap.error("с --invariants-only нужен OUT.docx (и, по желанию, IN.md перед ним)")
        md_path, docx_path = (None, args.files[0]) if len(args.files) == 1 else args.files
    else:
        if len(args.files) != 2:
            ap.error("нужны два файла: IN.md OUT.docx (или --invariants-only OUT.docx)")
        md_path, docx_path = args.files
    if args.expect_orient:
        try:
            parse_orient_list(args.expect_orient)
        except ValueError as e:
            ap.error(str(e))
    for pth in (md_path, docx_path):
        if pth and not Path(pth).is_file():
            print(f"Ошибка: нет файла {pth}", file=sys.stderr)
            return 1
    rep = run_checks(md_path, docx_path, args.expect_orient, args.landscape,
                     not args.no_auto_landscape, args.invariants_only)
    print(rep.render())
    return 0 if rep.ok else 1


if __name__ == "__main__":
    sys.exit(main())
