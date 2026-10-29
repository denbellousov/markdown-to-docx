# md2docx — Markdown to Word (DOCX) converter with clean tables and landscape pages

[![Release](https://img.shields.io/github/v/release/denbellousov/markdown-to-docx)](https://github.com/denbellousov/markdown-to-docx/releases)
![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-3776AB?logo=python&logoColor=white)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)

**English** | [Русский](README.ru.md)

**md2docx is a local Markdown to DOCX converter for documents with many tables.** It turns
a `.md` file into a Word document with Word's built-in heading styles, full-width tables whose
column widths are measured from the actual text, and mixed portrait and landscape pages:
a table too wide for a portrait page moves to a landscape page by itself, together with its
heading. md2docx is one Python file that needs no pandoc, no LibreOffice and no network,
plus an optional browser interface that runs on `127.0.0.1`.

![A portrait page with headings, a table and lists, and a landscape page with a wide eight-column table, both produced by md2docx from samples/08-showcase.md](docs/images/pages.png)

## Features

- **Headings** use Word's built-in Heading 1–6 styles, so the navigation pane and an automatic
  table of contents work.
- **Tables** span the full text width. Column widths come from the measured width of the text
  (Pillow and the real font files): short columns stay narrow, long text gets the room. The
  header row is grey and bold, borders are thin and grey, rows are not split between pages,
  and column alignment from `:---:` and `---:` is kept.
- **Automatic landscape pages.** A table that cannot fit a portrait page without breaking
  words moves to a landscape page together with its heading, caption and short introduction;
  the document returns to portrait right after it.
- **Explicit orientation and page breaks:** `<!-- landscape -->`, `<!-- portrait -->`,
  `<!-- pagebreak -->`, `\newpage`, `\pagebreak`. A `\newpage` between table rows splits the
  table at that row.
- **Readable on Windows.** The font table tells Word to substitute Arial for Helvetica Neue
  and Consolas for Menlo; `--fix-fonts` restores these hints after Word for Mac strips them.
- **Safe output.** Writes are atomic, and md2docx refuses to overwrite a document that was
  edited after conversion or is open in Word, unless you pass `--force`.
- **The rest of CommonMark, plus GFM strikethrough:** bold, italic, inline code, fenced and
  indented code blocks, nested quotes, nested lists with real Word numbering, hyperlinks, local
  images, horizontal rules, hard line breaks and `<br>`. YAML front matter and HTML comments are
  skipped; other HTML tags are removed and their text is kept, with a warning for HTML blocks.
- **Browser interface:** drop `.md` files or a whole folder, pick the orientation mode,
  download the `.docx` files one by one or as a `.zip`.

## Quick start

```sh
git clone https://github.com/denbellousov/markdown-to-docx.git
cd markdown-to-docx
pip install -r requirements.txt
python3 md2docx.py samples/08-showcase.md -o showcase.docx
```

Requirements: Python 3.10 or newer, `python-docx` and `markdown-it-py` (`lxml` comes with
`python-docx`). `Pillow` is optional but recommended: without it column widths are estimated
instead of measured.

md2docx is developed and tested on macOS with Python 3.13 (Word for Mac 16.114, LibreOffice
25.8.3.2). Column
widths are measured with Helvetica Neue and Arial from the macOS font folders or with Arial
from `C:\Windows\Fonts`; without these fonts they are estimated. With the Mac fonts hidden
(Arial only, and no fonts at all) every sample converts to the same page orientations and
passes the checker; md2docx has not been run on Windows or Linux itself. Command-line messages
and the browser interface are in Russian.

## Usage

```sh
python3 md2docx.py report.md                      # → report.docx next to the source
python3 md2docx.py report.md -o ~/Desktop/final.docx
python3 md2docx.py a.md b.md c.md                 # several files, each next to its source
python3 md2docx.py tables.md --landscape          # landscape by default (markers still switch)
python3 md2docx.py report.md --no-auto-landscape  # keep wide tables on portrait pages
python3 md2docx.py report.md --force              # overwrite the result even if it was edited
python3 md2docx.py --fix-fonts report.docx        # restore font hints after Word for Mac
python3 md2docx.py --version
```

| Option | What it does |
|---|---|
| `-o OUT.docx` | output path; only with a single input file |
| `--landscape` | the document starts in landscape; markers still switch the orientation |
| `--no-auto-landscape` | do not move wide tables to landscape pages |
| `--force` | overwrite an existing result even if it was edited after conversion, was not created by md2docx or is open in Word; with `--fix-fonts`, fix a file that is open in Word |
| `--fix-fonts FILE.docx …` | restore the Windows font hints in finished `.docx` files |
| `--version` | print the version |

For every file md2docx prints one summary line, for example
`Готово: report.docx — разделов 3 (К, А, К), таблиц 5, из них авто-альбомных 2`: done, three
sections (portrait, landscape, portrait), five tables, two of them moved to landscape
automatically. Warnings (a missing image, a link with an unsupported scheme, an HTML block,
a table that does not fit its page without breaking words) go to stderr with the prefix
`Внимание:`.
Exit codes: 0 — success; 1 — a file could not be read or written, including `-o` pointing
at the source (the source is not touched) and a result that must not be overwritten (see
[Overwrite protection](#overwrite-protection)); 2 — invalid arguments, including an empty
`-o`. An error in one file does not stop the rest of the batch.

Input must be UTF-8 (with or without BOM) or UTF-16 with BOM. Any other encoding is an error
(`файл не в кодировке UTF-8`, the file is not in UTF-8); null bytes without a BOM are reported
as `похоже на UTF-16 без BOM или двоичный файл — сохраните в UTF-8` (looks like UTF-16 without
BOM or a binary file — save it as UTF-8).

### Python API

```python
from md2docx import convert, fix_fonts

stats = convert("report.md", "report.docx", landscape=False, auto_landscape=True, force=False)
restored = fix_fonts("report.docx", force=False)
```

`convert()` returns a dict: `output` (the result path), `sections` (the section orientations
as letters, `К` for portrait and `А` for landscape, e.g. `['К', 'А', 'К']`), `tables`,
`auto_landscape_tables` and `warnings` (a list of warning lines, without the `Внимание:`
prefix). If the result must not be overwritten, it raises `FileExistsError` with the same
message as the CLI; `force=True` works like `--force`. It raises `OSError` (for example
`FileNotFoundError`) if a file cannot be read or written, and `UnicodeDecodeError` for an
unsupported encoding.

`fix_fonts()` does what `--fix-fonts` does and returns a dict "font → restored fields", empty
if the hints were already in place (the file is then not rewritten, and a Word lock does not
matter). It raises `FileExistsError` if the hints need restoring and the file is open in Word,
`ValueError` if it is not a valid `.docx`, and `OSError` if the file cannot be read or written.

## Page orientation and page breaks

Markers go on their own line between blocks. Case and spaces around the word do not matter:
`<!-- LANDSCAPE -->` is a marker, `<!-- land scape -->` is not.

| Marker | Effect |
|---|---|
| `<!-- landscape -->` | landscape pages from here on, starting on a new page |
| `<!-- portrait -->` | portrait pages from here on, starting on a new page |
| `<!-- pagebreak -->`, `\newpage`, `\pagebreak` | page break |

- A marker for the orientation already in effect does nothing, so no blank pages appear.
- A marker at the very beginning of the file sets the orientation of the first page, without
  a break.
- Several markers in a row: the orientation follows the last orientation marker, and there is
  a page break if any of them is a break; a new section starts on a new page anyway, so there
  is no second break. Markers at the end of the file are ignored. A block with no visible
  content (an empty quote, a comment, `<br>`, `&nbsp;`, an empty `#`) does not separate
  markers and does not get a page of its own.
- `\newpage` or `\pagebreak` right under the last row of a table (no blank line) is a page
  break after the table, not a table row. `\\newpage` with an escaped backslash and
  `&#92;newpage` are plain text.
- `\newpage` or `\pagebreak` between table rows splits the table at that row: the rest
  continues on a new page with the same columns and no header row; a table that moves to a
  landscape page moves as a whole. Several breaks give several parts, consecutive breaks count
  as one. A break right under the `|---|` line moves the whole table to a new page. A line
  without `|` under the break is already text after the table. An HTML comment
  `<!-- pagebreak -->` between table rows ends the table instead of splitting it (that is how
  GFM parses it): the rows below it become a paragraph with `|` characters. Use `\newpage`
  to split a table.
- Inside a code block a marker is plain text; inside a list or a quote it is ignored with
  a warning.

### Automatic landscape pages

A table that cannot fit a portrait page (17 cm of text width) without breaking words moves to
a landscape page by itself, and the document returns to portrait after it. A line may break
after a dash and after a hyphen, except a hyphen between a non-digit and a digit: `F-123456`
and `AB-2026` stay whole, as in LibreOffice. Dates, codes and phone numbers (`2026-10-22`,
`AB-2026-001`) stay whole in a table that fits a portrait page, but they never send a table
to landscape on their own.

What belongs to the table moves with it, so that a heading is not left alone at the bottom of
a portrait page:

- a caption paragraph ending with a colon right before the table;
- the heading above the table and all headings directly above it; between the heading and the
  table a short introduction is allowed — up to two paragraphs of at most 600 characters
  without images, and one horizontal rule;
- through the same kind of short introduction, a heading of a higher rank (a smaller number:
  `#` above `##`), so that a chapter moves together with a section that starts with a wide
  table.

Moving stops at a page break, the start of an explicitly marked section, a list, a quote,
code, another table, an HTML block with text, a long paragraph or a paragraph with an image,
a heading of the same or a lower rank (the same or a larger number), a third paragraph,
a second rule. Blocks with no visible content do not
interfere. Two wide tables with at most two visible blocks between them, no other tables and
no page break share one landscape section. `--no-auto-landscape` turns this off; sections
explicitly marked as landscape are never touched.

## Formatting

- A4, 2 cm margins on all sides in both orientations, no headers or footers.
- Body text: Helvetica Neue 11 pt, 8 pt after each paragraph, left-aligned.
- Headings: Word's built-in Heading 1–6 styles, bold, 18/16/14/12/11/11 pt, Heading 6 in
  italics.
- Tables: full text width, thin grey borders, a grey bold header row at the top of the table
  (it is not repeated on the following pages); a table whose header cells are all empty gets
  no header row. Rows are not split between pages, unless a row is taller than a page. A table
  does not start at the bottom of a page with its header and one row, and does not end with
  a single row on a new page, as long as those rows fit together: the start within ¾ of
  a page, the end within a page. A table at the end of the document does not produce a blank
  last page. Column widths follow the measured width of the text. If every cell fits on one
  line, each column gets its widest cell plus a share of the spare space in proportion to it;
  otherwise no column is narrower than its longest word, and the rest is shared in proportion
  to the amount of text in each column. A heading directly above a table is followed by 6 pt
  of space.
- Lists: real Word numbering; every list is numbered on its own and keeps its starting
  number. Indentation grows up to the eighth nesting level and stays there deeper.
- Quotes: grey italics with a line on the left. Code: Menlo 10 pt on a grey background with
  padding; adjacent code blocks stay separate.
- Links: hyperlinks only for `http:`, `https:`, `mailto:` and `tel:`. A link with another
  scheme (`file:`, `data:`, `smb:`, `javascript:` …) stays text, with a warning; a relative
  path or a `#anchor` stays text without a warning (there are no links inside the document).
- Footnotes (`[^1]` in the text and a `[^1]: …` line) stay as plain text in their places:
  there are no Word footnotes, but no text is lost.
- Images: local files are embedded, no wider than the text (or the cell) and no taller than
  0.8 of the text area height, keeping their proportions. A remote (URL), missing or
  unreadable image (for example SVG) and a file over 50 MB are not embedded: a placeholder
  `[изображение: …]` and a warning appear instead. An image outside the document's folder is
  embedded with a warning.
- Characters that Helvetica Neue lacks but Arial has (for example `→`) are set in Arial, so
  that Word for Mac does not draw them in Times New Roman. This needs Pillow and the macOS
  font files; elsewhere such characters stay in Helvetica Neue. Code is left as is.

### Known limitation: table rows at the bottom margin in LibreOffice

When the part of a table visible on a page ends exactly at the bottom margin, LibreOffice may
not render the following rows: they are in the file, but not in the PDF that LibreOffice
makes. This is a LibreOffice defect. Measured on LibreOffice 25.8.3.2 with a 7-row table and
the top margin shifted in 0.05 pt steps, 71 variants: in 10 consecutive variants (a 0.5 pt
window) the first page ended with row 2 right at the margin, and rows 3–7 were missing from
the PDF. Not repeating the header row does not fix it: md2docx 0.2.0, which repeated the
header, lost rows in the same 10 of 71 variants.

Word does not do this: in Word for Mac 16.114, 12 of 70 variants ended right at the margin
and no rows were lost, including with Arial metrics, as in Word for Windows without Helvetica
Neue (Word for Windows itself was not tested).

The workaround is `\newpage` between table rows above the missing ones, or before the table
(before its heading, if it has one). Everything below the break starts on a new page, and
LibreOffice renders it even if the part above the break still ends at the margin. On the same
71 offsets, a break after row 1, after row 2 or before the table heading lost no rows.

## Fonts on Windows and `--fix-fonts`

Windows usually has no Helvetica Neue. So that the document still looks the same, md2docx
writes substitution hints into the font table of the `.docx` (`fontTable.xml`):

- Helvetica Neue → `altName` Arial, family `swiss` (sans serif), `panose1`, variable pitch;
- Menlo (code) → `altName` Consolas, family `modern`, fixed pitch.

ISO/IEC 29500 (§17.8.3) defines these fields for choosing a substitute when a font is
missing; how Word uses `altName` is described in
[MS-OI29500 2.1.262](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/f7c44a60-3d38-4925-8934-57d77c277e84).
Without these hints Word would substitute a serif font.

The layout does not depend on the substitution. Some Arial glyphs are wider than Helvetica
Neue (the italic Cyrillic "т", brackets, quotes), so column widths use the larger of the two
measurements plus 3 %. Checked by rendering with Helvetica Neue forcibly replaced by Arial:
the page orientation stays the same and words in cells are not broken. There may be fewer
pages, since a line set in Arial is about 4 % lower. If Helvetica Neue is installed on the
reader's computer, it is used.

**After editing in Word for Mac, run `--fix-fonts`.** Every save in Word for Mac removes
`altName` from system fonts (MS-OI29500 2.1.262(c)) and also changes the Helvetica Neue family
from `swiss` to `auto`; Menlo loses `altName` Consolas and keeps `modern` (checked in Word for
Mac 16.114). Without these hints Word on Windows may substitute a serif font. `--fix-fonts`
(or `fix_fonts()` from Python) puts them back:

- it changes only the font table: the Helvetica Neue and Menlo entries get back the fields that
  md2docx writes, and if Helvetica Neue or Menlo is used in the text but has no entry in the
  table, the entry is added.
  The text and your edits are not touched, all other parts of the file keep the same bytes;
- run it after the last save in Word for Mac, before sending the file: saving it in Word for
  Mac again removes the hints again;
- a file open in Word (with a `~$…` lock file next to it) whose hints need restoring is not
  changed without `--force`: exit code 1, close the document and run again;
- the summary is `Готово: report.docx — восстановлены подсказки шрифтов: Helvetica Neue
  (altName, family), Menlo (altName)`; if the hints are in place, it is `Без изменений: …`
  and the file is not rewritten. Not a docx, a broken archive or no font table is an error
  for that file (exit code 1), and the other files are still fixed; the mode cannot be
  combined with `-o`, `--landscape` or `--no-auto-landscape` (exit code 2).

One limitation cannot be fixed inside the file: Word remembers a font substitution for the
session (MS-OI29500 2.1.261(a)). If a document with Helvetica Neue but without `altName` (for
example, one saved in Word for Mac) was opened earlier in the same session, your file gets the
same substitution. Quit Word completely and open the md2docx file first. A more reliable fix
is a permanent substitution on the reader's computer: File › Options › Advanced › Show
document content › Font Substitution, Helvetica Neue → Arial.

Also for Word on Windows: the document is marked as a Word 2013+ document
(`compatibilityMode` 15); otherwise Word opens it in compatibility mode and shifts tables into
the left margin.

## Overwrite protection

Every result is marked as an md2docx file (`dc:identifier` in the document properties; Word
does not show it). An existing file is overwritten silently only if it is an untouched md2docx
result: the mark is in place and the file has not been saved since the conversion (revision 1,
an empty "last modified by"). Otherwise md2docx refuses with exit code 1 and suggests
`--force`. Results of versions 0.1.0 and 0.1.1 have no mark, so only `--force` overwrites them.

A result that is open in Word is not replaced, even if it is untouched. Word puts a lock file
`~$…` next to an open file: for a name of up to 6 characters (without the extension) the lock
is `~$` plus the whole name (`report.docx` → `~$report.docx`), for 7 characters the first one
is dropped, for 8 or more the first two (`report-2026.docx` → `~$port-2026.docx`). If such
a lock exists, md2docx refuses with exit code 1 and the message `файл открыт в Word` (the file
is open in Word); `--force` overwrites in this case too. If Word is closed but the lock is still
there, delete the `~$…` file or use `--force`. Different files can share a lock name: an open
`cdefgh.docx` also blocks `abcdefgh.docx` (the lock of both is `~$cdefgh.docx`); `--force`
lifts it.

Writing is atomic: first a temporary file in the same folder, then a replace. If writing fails,
the previous result stays intact. A read-only result is not replaced even with `--force`
(error `нет доступа`, access denied). A symlinked result stays a symlink: the file it points to
is replaced. The permissions of the previous result are kept; hard links are not, since after
the replacement it is a new file. md2docx needs write permission for the output folder. Two
inputs of one run with the same output (`report.md` and `report.markdown` in one folder) are an
error for the later one, and the earlier result stays.

## Browser interface

![The md2docx browser interface after converting three sample files: the orientation switch, the result of each file and the download buttons](docs/images/web-ui.png)

To convert files without the command line:

```sh
python3 md2docx_web.py               # http://127.0.0.1:51536/ — opens the browser
python3 md2docx_web.py --no-browser  # only print the address
python3 md2docx_web.py --port N      # another port, if 51536 is busy
```

On macOS you can also double-click `md2docx-web.command` in Finder: a Terminal window with the
server opens. Stop the server with Ctrl+C in that window or by closing it.

Drop `.md` files or a whole folder onto the page (from a folder md2docx takes `.md`,
`.markdown` and images and skips hidden files and folders) or click «Выбрать файлы» (choose
files); the conversion starts right away. The orientation switch matches the CLI flags:
«Авто» (auto) — no flags, «Всё альбомное» (all landscape) — `--landscape`, «Без
автоальбомной» (no automatic landscape) — `--no-auto-landscape`; switching it re-converts the
same files. Each file shows «готово» (done) with the same summary line as the CLI, or «ошибка»
(error) with the message. A converted file also lists the converter's warnings verbatim and has
a «Скачать .docx» (download) button; when two or more files can be downloaded, «Скачать всё
(.zip)» appears as well. The result is the same as `md2docx.py` with the same flags.

A `.docx` dropped on the same page gets its fonts fixed for Windows, as with `--fix-fonts`.
It has to be a separate file: `.docx` files inside a dropped folder are skipped, since those
are usually earlier results. A copy is fixed; the original is not changed. Each `.docx` then
shows «шрифты исправлены» (fonts fixed) with what was restored, in the CLI's words, and
«Скачать .docx»; or «исправлять нечего — подсказки на месте» (nothing to fix, the hints are in
place), with nothing to download; or «ошибка» with the message. The orientation switch does not
apply to `.docx` files.

- The server listens on 127.0.0.1 only and answers only to `127.0.0.1` or `localhost`;
  API requests need a token that is generated at each server start and embedded in the page;
  after restarting the server, reload the page.
- The default port 51536 is fixed for the project; its pair 61536 is reserved too and stays
  unused, since there is no separate backend. If 51536 is busy, the server says so and exits;
  it never picks another port by itself.
- Source files are not changed: the browser sends copies, which live in the server's
  temporary folder. Only the last job is kept: a new selection deletes the copies and results
  of the previous one, and stopping the server (Ctrl+C, closing the window, `kill`) deletes
  everything. If the page is open in two tabs, a new selection in one makes the results in the
  other stale; that tab says so as soon as you use it (download or switch the orientation).
- Images are resolved relative to the `.md` file inside what you dropped: a path above the
  folder (`../`) is found only if you drop the common parent folder. Absolute paths and
  `file:` are read from disk, as in the CLI.
- Limits: 60 MB per file, 200 MB per selection.
- Names in the `.zip` do not repeat (case-insensitively): if a fixed `report.docx` and the
  result of `report.md` collide, the second one alphabetically gets a suffix,
  `report (2).docx`, as with `report.md` next to `report.markdown`. Each result on the page
  shows the name it downloads under on its own; in the `.zip`, files from a dropped folder keep
  their path from it: `folder/report.docx`.
- The Terminal window shows one log line per request (time, method, path, status), without
  file names or contents.

## FAQ

**How do I convert Markdown to Word with some pages in landscape?**
Put `<!-- landscape -->` on its own line before the part that should be landscape and
`<!-- portrait -->` after it, then run `python3 md2docx.py file.md`. Each marker starts a new
page in the new orientation. Wide tables need no markers: md2docx moves them to landscape
pages automatically.

**Why did my table end up on a landscape page?**
It did not fit the 17 cm text width of a portrait A4 page without breaking words. Run with
`--no-auto-landscape` to keep wide tables on portrait pages.

**Does md2docx need pandoc, LibreOffice or Microsoft Word?**
No. It needs Python 3.10+ with `python-docx`, `markdown-it-py` and, for measured column widths,
`Pillow`. LibreOffice and poppler are used only by the optional rendering tests.

**Will the document look right on Windows, where Helvetica Neue is missing?**
That is what the font table is for: it maps Helvetica Neue to Arial and Menlo to Consolas,
and column widths are computed for the wider of Helvetica Neue and Arial, so the substitution
does not break the layout. This was checked by rendering with Arial; Word for Windows itself
was not tested. If the file was re-saved in Word for Mac, run
`python3 md2docx.py --fix-fonts file.docx` before sending it.

**Is the table header repeated on every page?**
No. The grey header row appears once, at the top of the table. To start part of a long table
on a new page, put `\newpage` between its rows.

**Can md2docx overwrite a document I edited in Word?**
Not by accident: it refuses to overwrite a file that was saved after the conversion or is open
in Word. Pass `--force` when you mean it.

**Does it support footnotes and a table of contents?**
Footnotes stay as plain text where they are written. md2docx does not insert a table of
contents, but headings use Word's built-in styles, so Word can build one from them.

## Development and tests

```sh
pip install pytest
python3 -m pytest tests/ -q --basetemp=/tmp/md2docx-pytest -p no:cacheprovider
python3 tests/verify.py report.md report.docx     # check one document against its source
```

`--basetemp` keeps the temporary files of the tests in one known folder; use a dedicated one,
because pytest empties it at the start of a run. `-p no:cacheprovider` keeps `.pytest_cache`
out of the project.

`tests/verify.py` is an independent md ↔ docx checker: words, headings, tables, leftover markup
and the formatting invariants from [SPEC.md](SPEC.md). Check a document built with
`--landscape` or `--no-auto-landscape` with the same flag, otherwise the section orientations
count as a mismatch: `python3 tests/verify.py report.md report.docx --landscape`. Optional
parts of the test suite are skipped when their tools or data are missing:

- rendering to PDF needs LibreOffice (`MD2DOCX_SOFFICE` — the path to `soffice`,
  `MD2DOCX_LO_PROFILE` — its profile folder, temporary by default), poppler (`pdfinfo`,
  `pdftotext`, `pdffonts`, `pdftoppm`) and Pillow;
- regression on your own documents: list their paths in `tests/real_docs.local.txt` (ignored
  by git); the test deletes each converted result right after checking it.

The full behaviour contract is in [SPEC.md](SPEC.md) and the version history in
[CHANGELOG.md](CHANGELOG.md); both are in Russian.

## License

[MIT](LICENSE)
