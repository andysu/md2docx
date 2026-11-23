#!/usr/bin/env python3
"""
md2docx.py — універсальний конвертер Markdown у DOCX з підстановкою змінних.

ОДИН файл, без залежності від пакету oms_convert — усе включено:
  - модель AST-вузлів (dataclasses)
  - block-level парсер на markdown-it-py (таблиці, формули, виноски,
    task-list, definition list, front matter)
  - заповнення змінних через Jinja2 (front matter шаблону + YAML data-файли
    + --var з командного рядка), ДО парсингу markdown
  - DOCX-builder на python-docx з реальним вбудовуванням картинок
    (локальних і завантажених з http/https), формулами через Pandoc OMML
    (не math2docx — та бібліотека стабільно ламає \\sqrt) і власною
    системою нумерації списків (гарантований рестart на кожному
    незалежному списку — вбудовані стилі Word цього не гарантують)

ВСТАНОВЛЕННЯ ЗАЛЕЖНОСТЕЙ:
    pip install python-docx markdown-it-py mdit-py-plugins jinja2 pyyaml lxml pillow
    (потрібен також системний pandoc: apt install pandoc, або https://pandoc.org)

ВИКОРИСТАННЯ:
    # Просто конвертувати MD в DOCX (без змінних):
    python3 md2docx.py document.md -o document.docx

    # З підстановкою змінних із YAML-файлів (пізніші перекривають раніші):
    python3 md2docx.py template.md -o output.docx --data client.yaml --data station.yaml

    # З підстановкою окремих значень напряму з командного рядка:
    python3 md2docx.py template.md -o output.docx --var client.name="ТОВ Ромашка" --var act.number=17

    # Комбінація (--var має найвищий пріоритет, перекриває YAML і front matter):
    python3 md2docx.py template.md -o output.docx --data client.yaml --var act.number=42

    # Зберегти заповнений .md для перегляду перед конвертацією:
    python3 md2docx.py template.md -o output.docx --data client.yaml --keep-md filled.md

    # ПЕРЕВІРИТИ документ без конвертації (звіт: відсутні змінні/картинки/HTML):
    python3 md2docx.py document.md --check

    # Повний список параметрів:
    python3 md2docx.py --help

Синтаксис шаблону — звичайний Jinja2 усередині markdown: {{ client.name }},
{% for x in items %}...{% endfor %}, {% if cond %}...{% endif %}. Дефолтні
значення змінних можна задати у YAML front matter самого .md файлу —
--data і --var їх перекривають.
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import re
import subprocess
import sys
import tempfile
import urllib.request
import zipfile
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import Any

import docx
import yaml
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from jinja2 import Environment, StrictUndefined, UndefinedError
from lxml import etree
from markdown_it import MarkdownIt
from markdown_it.tree import SyntaxTreeNode
from mdit_py_plugins.deflist import deflist_plugin
from mdit_py_plugins.dollarmath import dollarmath_plugin
from mdit_py_plugins.footnote import footnote_plugin
from mdit_py_plugins.front_matter import front_matter_plugin
from mdit_py_plugins.tasklists import tasklists_plugin

# ============================================================
# ЧАСТИНА 1: AST-вузли
# ============================================================


from dataclasses import dataclass, field
from pathlib import Path


class Node:
    """Базовий маркерний клас усіх вузлів AST. Напряму не інстанціюється."""


# ============================================================
# Inline-level вузли — те, що може бути ВСЕРЕДИНІ Paragraph/Heading/TableCell
# ============================================================

@dataclass
class Text(Node):
    """Звичайний текстовий фрагмент без форматування."""
    value: str


@dataclass
class Bold(Node):
    children: list[Node]


@dataclass
class Italic(Node):
    children: list[Node]


@dataclass
class Strikethrough(Node):
    children: list[Node]


@dataclass
class InlineCode(Node):
    """Інлайн-код `так`. Без вкладених inline-вузлів — вміст береться дослівно."""
    value: str


@dataclass
class LineBreak(Node):
    """Розрив рядка всередині абзацу.
    hard=True  — жорсткий перенос (два пробіли в кінці рядка або <br>)
    hard=False — м'який перенос (просто новий рядок вихідного тексту)
    """
    hard: bool = True


@dataclass
class Link(Node):
    url: str
    title: str | None
    children: list[Node]


@dataclass
class AutoLink(Node):
    """<https://example.com> або <mailto:...> — посилання без окремого тексту."""
    url: str
    is_email: bool = False


@dataclass
class InlineImage(Node):
    """Картинка всередині рядка тексту (напр. клікабельна: [![alt](img.png)](url))."""
    src: str                       # шлях/URL як написано в MD
    resolved_path: Path | None     # абсолютний шлях на диску; None якщо це зовнішній URL
    alt: str
    title: str | None


@dataclass
class InlineMath(Node):
    """$E = mc^2$ — інлайн-формула, LaTeX-вираз без обгортки."""
    latex: str


@dataclass
class FootnoteRef(Node):
    """Посилання на виноску: текст[^1]."""
    identifier: str


@dataclass
class RawInlineHTML(Node):
    """Вбудований інлайн-HTML тег, напр. <u>...</u>, <sub>, <sup>."""
    html: str


# ============================================================
# Block-level вузли
# ============================================================

@dataclass
class Heading(Node):
    level: int              # 1..6
    children: list[Node]    # inline-вузли


@dataclass
class Paragraph(Node):
    children: list[Node]    # inline-вузли


@dataclass
class ThematicBreak(Node):
    """Горизонтальна лінія (---, ***, ___)."""


@dataclass
class CodeBlock(Node):
    code: str
    language: str | None    # напр. "python", "json", "bash"; None якщо мова не вказана


@dataclass
class MathBlock(Node):
    """$$...$$ — блочна формула на весь рядок/абзац."""
    latex: str


@dataclass
class ListItem(Node):
    children: list[Node]         # block-level вузли (Paragraph, вкладений List, Table...)
    checked: bool | None = None  # None = звичайний пункт; True/False = task-list чекбокс


@dataclass
class List(Node):
    ordered: bool
    items: list[ListItem]
    start: int = 1                # з якого номера починається нумерований список


@dataclass
class DefinitionItem(Node):
    term: list[Node]                    # inline-вузли терміна
    definitions: list[list[Node]]       # кожен опис — свій список inline-вузлів


@dataclass
class DefinitionList(Node):
    items: list[DefinitionItem]


@dataclass
class Blockquote(Node):
    """Цитата. children — БЛОЧНІ вузли, рекурсивно (список, таблиця, код,
    навіть вкладена цитата — секції 4 і 14 тестового файлу саме це перевіряють)."""
    children: list[Node]


@dataclass
class TableCell(Node):
    children: list[Node]              # inline-вузли
    align: str | None = None          # "left" | "center" | "right" | None


@dataclass
class TableRow(Node):
    cells: list[TableCell]
    is_header: bool = False


@dataclass
class Table(Node):
    rows: list[TableRow]              # перший рядок з is_header=True — заголовок


@dataclass
class Image(Node):
    """Картинка як окремий блок (не всередині речення)."""
    src: str
    resolved_path: Path | None
    alt: str
    title: str | None
    is_link_wrapped: bool = False     # чи обгорнута в [![alt](img)](url)
    link_url: str | None = None


@dataclass
class Callout(Node):
    """Блок інформаційного виділення / GitHub Alert (> [!NOTE], > [!WARNING] тощо)."""
    alert_type: str                  # "NOTE" | "TIP" | "IMPORTANT" | "WARNING" | "CAUTION"
    children: list[Node] = field(default_factory=list)


@dataclass
class MermaidDiagram(Node):
    """Mermaid діаграма (```mermaid ... ```) для авторендерингу в вектор/растр."""
    code: str
    alt: str = "Mermaid Diagram"


@dataclass
class HTMLBlock(Node):
    """Блок вбудованого HTML: <div>...</div>, <details>...</details> тощо."""
    html: str


@dataclass
class FootnoteDefinition(Node):
    identifier: str
    children: list[Node]              # block-level (текст виноски може мати кілька абзаців)


@dataclass
class Document(Node):
    """Корінь дерева одного MD-файлу."""
    children: list[Node] = field(default_factory=list)          # block-level вузли верхнього рівня
    footnotes: dict[str, FootnoteDefinition] = field(default_factory=dict)
    source_path: Path | None = None    # каталог цього файлу — база для резолву відносних шляхів картинок
    front_matter: dict = field(default_factory=dict)             # YAML front matter (метадані, не рендериться)


# ============================================================
# ЧАСТИНА 2: LaTeX -> OMML (формули через Pandoc)
# ============================================================


import subprocess
import tempfile
import zipfile
from functools import lru_cache
from pathlib import Path

from lxml import etree

NS_M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


class MathConversionError(Exception):
    """LaTeX не вдалось конвертувати в OMML (помилка Pandoc або формула не знайдена в результаті)."""


@lru_cache(maxsize=1024)
def _latex_to_omml_xml(latex: str, display: bool) -> bytes:
    delim = "$$" if display else "$"
    md_text = f"{delim}{latex}{delim}\n"

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        md_file = tmp_path / "formula.md"
        docx_file = tmp_path / "formula.docx"
        md_file.write_text(md_text, encoding="utf-8")

        result = subprocess.run(
            ["pandoc", str(md_file), "-o", str(docx_file)],
            capture_output=True, timeout=30,
        )
        if result.returncode != 0:
            raise MathConversionError(
                f"Pandoc не зміг сконвертувати формулу {latex!r}: "
                f"{result.stderr.decode('utf-8', 'ignore')}"
            )

        with zipfile.ZipFile(docx_file) as z:
            xml_bytes = z.read("word/document.xml")

    tree = etree.fromstring(xml_bytes)
    tag = "oMathPara" if display else "oMath"
    node = tree.find(f".//{NS_M}{tag}")
    if node is None:
        node = tree.find(f".//{NS_M}oMath")  # fallback, якщо Pandoc не обгорнув у oMathPara
    if node is None:
        raise MathConversionError(f"Pandoc не згенерував OMML для формули: {latex!r}")

    return etree.tostring(node)


def insert_math(paragraph, latex: str, display: bool = False) -> None:
    """Вставляє формулу в кінець python-docx Paragraph як справжній OMML-об'єкт Word
    (не картинку і не текст) — редагований у Word так само, як формула, вставлена вручну.

    Викликати в тому самому порядку, що й paragraph.add_run() для сусіднього тексту —
    OMML-вузол додається в кінець XML параграфа, тож послідовність викликів визначає
    порядок вмісту в рядку.
    """
    if not latex.strip():
        return  # порожня формула (напр. артефакт неоднозначного джерела) — тихо пропускаємо
    xml_bytes = _latex_to_omml_xml(latex, display)
    node = etree.fromstring(xml_bytes)
    paragraph._p.append(node)


# ============================================================
# ЧАСТИНА 3: Власна нумерація списків
# ============================================================


from docx.oxml import OxmlElement
from docx.oxml.ns import qn

_BULLET_CHARS = ["\u2022", "\u25E6", "\u25AA", "\u2023"]  # •, ◦, ▪, ‣
_ABSTRACT_BULLET_ID = 9001
_ABSTRACT_DECIMAL_ID = 9002
_NEXT_NUM_ID_ATTR = "_oms_next_num_id"


def _numbering_element(container):
    return container.part.numbering_part.element


def _ensure_abstract_nums(container) -> None:
    """Реєструє наші власні abstractNum-и в numbering.xml, якщо їх ще нема
    (ідемпотентно — безпечно викликати перед кожним списком)."""
    numbering = _numbering_element(container)
    if numbering.find(f'{qn("w:abstractNum")}[@{qn("w:abstractNumId")}="{_ABSTRACT_BULLET_ID}"]') is not None:
        return  # вже зареєстровано раніше в цьому документі

    numbering.insert(0, _build_abstract_num(_ABSTRACT_DECIMAL_ID, ordered=True))
    numbering.insert(0, _build_abstract_num(_ABSTRACT_BULLET_ID, ordered=False))


def _build_abstract_num(abstract_id: int, ordered: bool):
    abstract_num = OxmlElement("w:abstractNum")
    abstract_num.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "multilevel")
    abstract_num.append(multi)

    for ilvl in range(4):
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(ilvl))

        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")
        lvl.append(start)

        num_fmt = OxmlElement("w:numFmt")
        lvl_text = OxmlElement("w:lvlText")
        if ordered:
            num_fmt.set(qn("w:val"), "decimal")
            lvl_text.set(qn("w:val"), f"%{ilvl + 1}.")
        else:
            num_fmt.set(qn("w:val"), "bullet")
            lvl_text.set(qn("w:val"), _BULLET_CHARS[ilvl % len(_BULLET_CHARS)])
        lvl.append(num_fmt)
        lvl.append(lvl_text)

        jc = OxmlElement("w:lvlJc")
        jc.set(qn("w:val"), "left")
        lvl.append(jc)

        pPr = OxmlElement("w:pPr")
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), str(360 + ilvl * 360))
        ind.set(qn("w:hanging"), "360")
        pPr.append(ind)
        lvl.append(pPr)

        rPr = OxmlElement("w:rPr")
        if not ordered:
            rFonts = OxmlElement("w:rFonts")
            rFonts.set(qn("w:ascii"), "Arial")
            rFonts.set(qn("w:hAnsi"), "Arial")
            rFonts.set(qn("w:hint"), "default")
            rPr.append(rFonts)
        lvl.append(rPr)

        abstract_num.append(lvl)

    return abstract_num


def new_list_numid(container, ordered: bool) -> int:
    """Створює нову незалежну інстанцію нумерації (restart із 1) і повертає її numId.
    Викликати РІВНО ОДИН РАЗ на кожен незалежний список верхнього рівня —
    вкладені під-списки мають переюзати цей самий numId з іншим ilvl.

    container — Document АБО _Cell (обидва мають .part, що резолвиться до
    того самого пакета документа — лічильник і numbering.xml завжди спільні)."""
    _ensure_abstract_nums(container)
    part = container.part
    numbering = _numbering_element(container)

    next_id = getattr(part, _NEXT_NUM_ID_ATTR, None)
    if next_id is None:
        existing = numbering.findall(qn("w:num"))
        next_id = 9101 + len(existing)
    setattr(part, _NEXT_NUM_ID_ATTR, next_id + 1)

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(next_id))
    abstract_ref = OxmlElement("w:abstractNumId")
    abstract_ref.set(qn("w:val"), str(_ABSTRACT_DECIMAL_ID if ordered else _ABSTRACT_BULLET_ID))
    num.append(abstract_ref)
    # КЛЮЧОВЕ: сам факт іншого numId НЕ гарантує рестарт лічильника з 1 —
    # емпірично перевірено (LibreOffice/Word продовжують лік, бо посилаються
    # на той самий abstractNumId). Явний startOverride на кожному рівні —
    # єдиний надійний спосіб примусового рестарту для незалежного списку.
    for ilvl in range(4):
        lvl_override = OxmlElement("w:lvlOverride")
        lvl_override.set(qn("w:ilvl"), str(ilvl))
        start_override = OxmlElement("w:startOverride")
        start_override.set(qn("w:val"), "1")
        lvl_override.append(start_override)
        num.append(lvl_override)
    numbering.append(num)

    return next_id


def apply_numbering(paragraph, num_id: int, ilvl: int) -> None:
    """Прив'язує абзац до конкретної інстанції нумерації на конкретному рівні
    вкладеності — напряму через w:numPr, в обхід іменованих стилів списку."""
    pPr = paragraph._p.get_or_add_pPr()
    numPr = OxmlElement("w:numPr")
    ilvl_el = OxmlElement("w:ilvl")
    ilvl_el.set(qn("w:val"), str(min(ilvl, 3)))
    num_id_el = OxmlElement("w:numId")
    num_id_el.set(qn("w:val"), str(num_id))
    numPr.append(ilvl_el)
    numPr.append(num_id_el)
    pPr.insert(0, numPr)  # numPr — один з перших елементів у схемі CT_PPrBase


# ============================================================
# ЧАСТИНА 4: Block-level парсер markdown -> AST
# ============================================================


import html as html_module
import re
from pathlib import Path

from markdown_it import MarkdownIt
from markdown_it.tree import SyntaxTreeNode
from mdit_py_plugins.deflist import deflist_plugin
from mdit_py_plugins.dollarmath import dollarmath_plugin
from mdit_py_plugins.footnote import footnote_plugin
from mdit_py_plugins.front_matter import front_matter_plugin
from mdit_py_plugins.tasklists import tasklists_plugin


_EXTERNAL_URL_RE = re.compile(r"^([a-zA-Z][a-zA-Z0-9+.\-]*:)?//|^data:", re.IGNORECASE)
# ловить http://, https://, //cdn..., data:image/... — усе, що НЕ треба
# резолвити як локальний файл на диску.


def _build_md() -> MarkdownIt:
    return (
        MarkdownIt("gfm-like")
        .use(front_matter_plugin)
        .use(footnote_plugin)
        .use(deflist_plugin)
        .use(tasklists_plugin)
        .use(dollarmath_plugin)
    )


_MD = _build_md()

_STANDALONE_BLOCK_MATH_RE = re.compile(r"^([ \t]*)\$\$(.+)\$\$([ \t]*)$")


def _ensure_block_math_isolation(text: str) -> str:
    """Якщо рядок повністю складається з $$...$$ (блочна формула), але йде
    одразу після непорожнього рядка без порожнього рядка між ними — типовий
    патерн у нумерованих списках ("Текст:\\n   $$formula$$") — CommonMark
    трактує це як lazy continuation того самого абзацу: markdown-it більше
    НЕ розпізнає $$ як окрему блочну формулу, і вона перетворюється на inline
    зі сторонніми символами $ навколо. Вставляємо порожній рядок перед і
    (за потреби) після такої формули, зберігаючи відступ рядка незмінним —
    відступ визначає приналежність до списку, і вставлений порожній рядок
    сам по собі приналежність не рве (список просто стає "loose")."""
    lines = text.split("\n")
    out: list[str] = []
    n = len(lines)
    for i, line in enumerate(lines):
        is_block_math = bool(_STANDALONE_BLOCK_MATH_RE.match(line))
        if is_block_math and out and out[-1].strip() != "":
            out.append("")
        out.append(line)
        if is_block_math and i + 1 < n and lines[i + 1].strip() != "":
            out.append("")
    return "\n".join(out)


# ============================================================
# Точка входу
# ============================================================

def parse_markdown(text: str, source_path: Path | None = None) -> Document:
    """Парсить текст markdown у Document AST.

    source_path — шлях до .md файлу (або його каталогу), потрібен для
    резолву відносних шляхів картинок (images/foo.png -> абсолютний Path).
    """
    tokens = _MD.parse(_ensure_block_math_isolation(text))
    tree = SyntaxTreeNode(tokens)
    base_dir = _base_dir(source_path)

    document = Document(source_path=source_path)

    for child in tree.children:
        if child.type == "front_matter":
            document.front_matter = _parse_front_matter(child.content)
            continue
        if child.type == "footnote_block":
            for fn in child.children:
                fn_node = _convert_footnote(fn, base_dir)
                document.footnotes[fn_node.identifier] = fn_node
            continue
        blocks = _convert_block(child, base_dir)
        document.children.extend(blocks)

    return document


def _base_dir(source_path: Path | None) -> Path | None:
    if source_path is None:
        return None
    return source_path.parent if source_path.suffix else source_path


def _parse_front_matter(raw: str) -> dict:
    import yaml
    data = yaml.safe_load(raw) or {}
    return data if isinstance(data, dict) else {}


def _resolve_image_path(src: str, base_dir: Path | None) -> Path | None:
    """Абсолютний шлях до картинки на диску, якщо src — локальний файл.
    Повертає None для зовнішніх URL (http://, data:...) — такі конвертер
    пізніше або довантажить окремо, або пропустить з попередженням."""
    if _EXTERNAL_URL_RE.match(src):
        return None
    if base_dir is None:
        return Path(src)
    return (base_dir / src).resolve()


# ============================================================
# Block-level конвертація
# ============================================================

def _convert_block(node: SyntaxTreeNode, base_dir: Path | None) -> list[Node]:
    t = node.type

    if t == "heading":
        level = int(node.tag[1])  # 'h3' -> 3
        return [Heading(level=level, children=_convert_inline(node, base_dir))]

    if t == "paragraph":
        image_block = _paragraph_as_image(node, base_dir)
        if image_block is not None:
            return [image_block]
        return [Paragraph(children=_convert_inline(node, base_dir))]

    if t == "hr":
        return [ThematicBreak()]

    if t == "fence" or t == "code_block":
        lang = (node.info.strip().lower() or None) if t == "fence" else None
        if lang == "mermaid":
            code_content = node.content.strip()
            if code_content:
                return [MermaidDiagram(code=code_content)]
        return [CodeBlock(code=node.content, language=(node.info.strip() or None) if t == "fence" else None)]

    if t == "math_block":
        latex = node.content.strip()
        return [MathBlock(latex=latex)] if latex else []

    if t == "html_block":
        return _convert_html_block(node.content, base_dir)

    if t == "blockquote":
        children = [c for ch in node.children for c in _convert_block(ch, base_dir)]
        callout = _detect_callout(children)
        if callout is not None:
            return [callout]
        return [Blockquote(children=children)]

    if t == "bullet_list":
        return [_convert_list(node, ordered=False, base_dir=base_dir)]

    if t == "ordered_list":
        return [_convert_list(node, ordered=True, base_dir=base_dir)]

    if t == "table":
        return [_convert_table(node, base_dir)]

    if t == "dl":
        return [_convert_deflist(node, base_dir)]

    if t == "footnote_block":
        # верхнього рівня обробляється окремо в parse_markdown; якщо трапився
        # вкладено (нетипово) — просто ігноруємо, виноски тут не рендеряться інлайново.
        return []

    # Невідомий/нерелевантний тип вузла — не валимо весь парсинг,
    # а сигналізуємо явним "непідтримувано", щоб не загубити контент мовчки.
    return [HTMLBlock(html=f"<!-- НЕПІДТРИМУВАНИЙ БЛОК: {t} -->")]


_IMG_TAG_RE = re.compile(r"<img\b[^>]*>", re.IGNORECASE)
_ATTR_RE = re.compile(r'(\w[\w-]*)\s*=\s*"([^"]*)"')


def _convert_html_block(html: str, base_dir: Path | None) -> list[Node]:
    """Сирий HTML-блок (<div>, <img>, коментар...) markdown сам не розбирає,
    а Word не виконує довільний HTML — тому за замовчуванням текст усередині
    зберігається (builder дістає з нього голий текст на випадок непорожнього
    вмісту). АЛЕ <img> усередині такого блоку — дуже поширений реальний кейс
    (книги/сайти, експортовані в markdown, де картинки вставлені як
    <div class="book-figure"><img src=".."/><div class="figure-caption">..</div></div>,
    а не через ![]()) — і його варто розпізнавати спеціально, інакше факт
    існування зображення повністю губиться. Тому blок ділиться на послідовність:
    [текст до першого img] [Image] [текст між img] [Image] ... [текст після] —
    кожен текстовий шматок іде як HTMLBlock (спрощується в builder), а кожен
    <img> — як повноцінний Image з резолвленим шляхом (той самий механізм,
    що й для markdown-картинок)."""
    matches = list(_IMG_TAG_RE.finditer(html))
    if not matches:
        return [HTMLBlock(html=html)]

    nodes: list[Node] = []
    pos = 0
    for m in matches:
        before = html[pos:m.start()]
        if before.strip():
            nodes.append(HTMLBlock(html=before))
        attrs = dict(_ATTR_RE.findall(m.group(0)))
        src = html_module.unescape(attrs.get("src", ""))
        if src:
            nodes.append(Image(
                src=src,
                resolved_path=_resolve_image_path(src, base_dir),
                alt=html_module.unescape(attrs.get("alt", "")),
                title=html_module.unescape(attrs.get("title")) if attrs.get("title") else None,
            ))
        pos = m.end()
    tail = html[pos:]
    if tail.strip():
        nodes.append(HTMLBlock(html=tail))
    return nodes


_CALLOUT_RE = re.compile(r"^\s*\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\][ \t]*", re.IGNORECASE)


def _detect_callout(children: list[Node]) -> Callout | None:
    """Розпізнає GitHub Alerts / Callout-блоки у цитатах:
    > [!NOTE] / > [!WARNING] тощо.
    Вилучає технічний маркер [!TYPE] і повертає вузол Callout.
    """
    if not children:
        return None
    first_node = children[0]
    if not isinstance(first_node, Paragraph) or not first_node.children:
        return None
    first_inline = first_node.children[0]
    if not isinstance(first_inline, Text):
        return None
    m = _CALLOUT_RE.match(first_inline.value)
    if not m:
        return None
    alert_type = m.group(1).upper()
    rem_text = first_inline.value[m.end():]
    new_p_children = list(first_node.children)
    if rem_text.strip():
        new_p_children[0] = Text(value=rem_text.lstrip())
    else:
        new_p_children = new_p_children[1:]
        while new_p_children and isinstance(new_p_children[0], LineBreak):
            new_p_children = new_p_children[1:]

    new_children: list[Node] = []
    if new_p_children:
        new_children.append(Paragraph(children=new_p_children))
    new_children.extend(children[1:])
    return Callout(alert_type=alert_type, children=new_children)


def _paragraph_as_image(node: SyntaxTreeNode, base_dir: Path | None) -> Image | None:
    """Абзац, що складається ТІЛЬКИ з картинки (можливо, обгорнутої в посилання) —
    трактуємо як окремий block-level Image, а не Paragraph[InlineImage].
    Саме так markdown-it представляє '![alt](img.png)' на своєму рядку."""
    inline = node.children[0] if node.children else None
    if inline is None or inline.type != "inline":
        return None
    kids = inline.children
    if len(kids) != 1:
        return None
    only = kids[0]

    if only.type == "image":
        return _image_from_node(only, base_dir, is_link_wrapped=False, link_url=None)

    if only.type == "link" and len(only.children) == 1 and only.children[0].type == "image":
        return _image_from_node(
            only.children[0], base_dir, is_link_wrapped=True, link_url=only.attrs.get("href")
        )

    return None


def _image_from_node(
    node: SyntaxTreeNode, base_dir: Path | None, is_link_wrapped: bool, link_url: str | None
) -> Image:
    src = node.attrs.get("src", "")
    return Image(
        src=src,
        resolved_path=_resolve_image_path(src, base_dir),
        alt=_plain_text(node.children),
        title=node.attrs.get("title") or None,
        is_link_wrapped=is_link_wrapped,
        link_url=link_url,
    )


def _convert_list(node: SyntaxTreeNode, ordered: bool, base_dir: Path | None) -> List:
    start = int(node.attrs.get("start", 1)) if ordered else 1
    items = [_convert_list_item(child, base_dir) for child in node.children]
    return List(ordered=ordered, items=items, start=start)


def _convert_list_item(node: SyntaxTreeNode, base_dir: Path | None) -> ListItem:
    is_task = "task-list-item" in node.attrs.get("class", "")
    checked: bool | None = None

    children: list[Node] = []
    for i, child in enumerate(node.children):
        blocks = _convert_block(child, base_dir)
        for j, block in enumerate(blocks):
            if is_task and i == 0 and j == 0 and isinstance(block, Paragraph):
                checked, block = _extract_checkbox(block)
            children.append(block)

    return ListItem(children=children, checked=checked)


def _extract_checkbox(paragraph: Paragraph) -> tuple[bool, Paragraph]:
    """Витягує чекбокс task-list зі старту абзацу (вставлений плагіном tasklists
    як html_inline <input ... checked?>) і повертає (checked, абзац_без_чекбокса)."""
    if paragraph.children and isinstance(paragraph.children[0], RawInlineHTML):
        checked = "checked" in paragraph.children[0].html
        remaining = paragraph.children[1:]
        # прибираємо можливий пробіл одразу після маркера
        if remaining and isinstance(remaining[0], Text):
            remaining[0] = Text(value=remaining[0].value.lstrip())
        return checked, Paragraph(children=remaining)
    return False, paragraph


def _convert_table(node: SyntaxTreeNode, base_dir: Path | None) -> Table:
    rows: list[TableRow] = []
    for section in node.children:  # thead / tbody
        is_header_section = section.type == "thead"
        for tr in section.children:
            cells = [_convert_table_cell(cell, base_dir) for cell in tr.children]
            rows.append(TableRow(cells=cells, is_header=is_header_section))
    return Table(rows=rows)


def _convert_table_cell(node: SyntaxTreeNode, base_dir: Path | None) -> TableCell:
    align = None
    style = node.attrs.get("style", "")
    m = re.search(r"text-align:\s*(\w+)", style)
    if m:
        align = m.group(1)
    return TableCell(children=_convert_inline(node, base_dir), align=align)


def _convert_deflist(node: SyntaxTreeNode, base_dir: Path | None) -> DefinitionList:
    items: list[DefinitionItem] = []
    term: list[Node] | None = None
    definitions: list[list[Node]] = []

    def flush() -> None:
        if term is not None:
            items.append(DefinitionItem(term=term, definitions=definitions))

    for child in node.children:
        if child.type == "dt":
            if term is not None:
                flush()
            term = _convert_inline(child, base_dir)
            definitions = []
        elif child.type == "dd":
            # dd містить block-level вміст (зазвичай paragraph) — беремо inline з нього
            inline_children: list[Node] = []
            for sub in child.children:
                if sub.type == "paragraph":
                    inline_children.extend(_convert_inline(sub, base_dir))
                else:
                    inline_children.extend(_plain_text_as_nodes(sub))
            definitions.append(inline_children)
    flush()

    return DefinitionList(items=items)


def _plain_text_as_nodes(node: SyntaxTreeNode) -> list[Node]:
    return [Text(value=_plain_text([node]))]


def _convert_footnote(node: SyntaxTreeNode, base_dir: Path | None) -> FootnoteDefinition:
    identifier = str(node.meta.get("label"))
    children: list[Node] = []
    for child in node.children:
        blocks = _convert_block(child, base_dir)
        children.extend(_strip_footnote_anchor(b) for b in blocks)
    return FootnoteDefinition(identifier=identifier, children=children)


def _strip_footnote_anchor(block: Node) -> Node:
    """Видаляє FootnoteRef-подібний backlink-маркер (↩), який плагін footnote
    додає в кінець останнього абзацу визначення виноски — він не потрібен
    у друкованому документі."""
    if isinstance(block, Paragraph) and block.children:
        last = block.children[-1]
        if isinstance(last, RawInlineHTML) or (isinstance(last, Text) and last.value.strip() == "↩"):
            block.children = block.children[:-1]
    return block


# ============================================================
# Inline-level конвертація
# ============================================================

def _convert_inline(node: SyntaxTreeNode, base_dir: Path | None) -> list[Node]:
    """Бере block-вузол (heading/paragraph/table_cell/...), знаходить його
    єдину дочірню 'inline' ноду і рекурсивно перетворює її дітей."""
    if not node.children:
        return []
    inline = node.children[0]
    if inline.type != "inline":
        return []
    return _convert_inline_list(inline.children, base_dir)


def _convert_inline_list(nodes: list[SyntaxTreeNode], base_dir: Path | None) -> list[Node]:
    result: list[Node] = []
    for n in nodes:
        converted = _convert_inline_node(n, base_dir)
        if converted is not None:
            result.append(converted)
    return result


def _convert_inline_node(node: SyntaxTreeNode, base_dir: Path | None) -> Node | None:
    t = node.type

    if t == "text":
        return Text(value=node.content)
    if t == "softbreak":
        return LineBreak(hard=False)
    if t == "hardbreak":
        return LineBreak(hard=True)
    if t == "code_inline":
        return InlineCode(value=node.content)
    if t == "math_inline":
        latex = node.content.strip()
        if not latex:
            # Порожня формула — майже завжди артефакт неоднозначного джерела
            # (напр. два блочні $$...$$ на continuation-рядках цитати без порожнього
            # рядка між ними зливаються в один inline-потік, і dollarmath видає
            # "порожній" $$ posередині). Пропускаємо вузол, а не падаємо на порожньому
            # LaTeX — інакше одна ця формула валить конвертацію всього документа.
            return None
        return InlineMath(latex=latex)
    if t == "strong":
        return Bold(children=_convert_inline_list(node.children, base_dir))
    if t == "em":
        return Italic(children=_convert_inline_list(node.children, base_dir))
    if t == "s":
        return Strikethrough(children=_convert_inline_list(node.children, base_dir))
    if t == "html_inline":
        return RawInlineHTML(html=node.content)
    if t == "footnote_ref":
        return FootnoteRef(identifier=str(node.meta.get("label")))
    if t == "footnote_anchor":
        return None  # backlink-стрілка ↩, не рендеримо

    if t == "link":
        href = node.attrs.get("href", "")
        if node.markup in ("autolink", "linkify"):
            return AutoLink(url=href, is_email=href.startswith("mailto:"))
        return Link(
            url=href,
            title=node.attrs.get("title") or None,
            children=_convert_inline_list(node.children, base_dir),
        )

    if t == "image":
        src = node.attrs.get("src", "")
        return InlineImage(
            src=src,
            resolved_path=_resolve_image_path(src, base_dir),
            alt=_plain_text(node.children),
            title=node.attrs.get("title") or None,
        )

    # Невідомий inline-тип — не втрачаємо текст мовчки.
    return Text(value=f"[НЕПІДТРИМУВАНИЙ INLINE: {t}]")


def _plain_text(nodes: list[SyntaxTreeNode]) -> str:
    """Згортає inline-вузли (напр. alt-текст картинки з **bold** усередині)
    у звичайний рядок — Word alt-text не підтримує форматування."""
    parts: list[str] = []
    for n in nodes:
        if n.type == "text":
            parts.append(n.content)
        elif n.children:
            parts.append(_plain_text(n.children))
    return "".join(parts)


# ============================================================
# ЧАСТИНА 5: Заповнення змінних (Jinja2 + YAML)
# ============================================================


import argparse
import re
import sys
from pathlib import Path
from typing import Any

import yaml
from jinja2 import Environment, StrictUndefined, Undefined, UndefinedError

FRONT_MATTER_RE = re.compile(r"\A---\s*\n(.*?\n)---\s*\n", re.DOTALL)


def split_front_matter(text: str) -> tuple[dict[str, Any], str]:
    """Розділяє MD-файл на YAML front matter (dict) і тіло документа.

    Front matter слугує значеннями ЗА ЗАМОВЧУВАННЯМ для змінних шаблону
    (напр. назва компанії-виконавця), які зовнішній data-файл може перевизначити.
    """
    match = FRONT_MATTER_RE.match(text)
    if not match:
        return {}, text
    raw_front_matter = match.group(1)
    body = text[match.end():]
    data = yaml.safe_load(raw_front_matter) or {}
    if not isinstance(data, dict):
        raise ValueError("Front matter має бути YAML-словником (key: value)")
    return data, body


def load_data_files(paths: list[Path]) -> dict[str, Any]:
    """Завантажує й зливає кілька YAML data-файлів по порядку.

    Пізніші файли перевизначають значення з попередніх — зручно для схеми
    "базові реквізити компанії" + "дані конкретного клієнта" в окремих файлах.
    """
    merged: dict[str, Any] = {}
    for path in paths:
        if not path.exists():
            raise FileNotFoundError(f"Data-файл не знайдено: {path}")
        with path.open("r", encoding="utf-8") as f:
            chunk = yaml.safe_load(f) or {}
        if not isinstance(chunk, dict):
            raise ValueError(f"{path} має містити YAML-словник верхнього рівня")
        merged = _deep_merge(merged, chunk)
    return merged


def _deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    """Рекурсивне злиття словників — дозволяє перевизначати лише окреме вкладене
    поле (напр. client.address), не чіпаючи інші поля client."""
    result = dict(base)
    for key, value in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


class MissingVariableError(Exception):
    """Кидається, коли шаблон посилається на змінну, якої немає в даних.

    Навмисно fail-fast: краще впасти з чіткою помилкою на етапі заповнення,
    ніж надрукувати документ з дірою на місці реквізитів.
    """


def render_template(
    template_path: Path,
    data_paths: list[Path],
    output_path: Path | None = None,
) -> str:
    """Заповнює MD-шаблон значеннями і повертає готовий текст.

    Порядок джерел змінних (пізніше перекриває раніше):
      1. front matter самого шаблону (дефолти)
      2. data-файли в переданому порядку

    Якщо output_path задано — записує результат туди (без front matter:
    метадані вже "розчинені" в тексті документа, front matter далі не потрібен).
    """
    raw_text = template_path.read_text(encoding="utf-8")
    front_matter_defaults, body = split_front_matter(raw_text)
    external_data = load_data_files(data_paths)
    context = _deep_merge(front_matter_defaults, external_data)

    env = Environment(
        undefined=StrictUndefined,  # падає одразу на неіснуючій змінній, не мовчки друкує "Undefined"
        autoescape=False,           # це не HTML — не треба екранувати & < > "
        trim_blocks=True,
        lstrip_blocks=True,
    )
    template = env.from_string(body)

    try:
        rendered = template.render(**context)
    except UndefinedError as exc:
        raise MissingVariableError(
            f"Шаблон '{template_path.name}' посилається на змінну, "
            f"якої немає в даних: {exc}"
        ) from exc

    if output_path is not None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(rendered, encoding="utf-8")

    return rendered


def render_template_collecting(
    template_path: Path,
    data_paths: list[Path],
    cli_vars: dict[str, Any] | None = None,
) -> tuple[str, list[str]]:
    """Те саме заповнення, що й render_template, але НІКОЛИ не падає на
    відсутній змінній — замість цього збирає імена ВСІХ відсутніх змінних
    (а не тільки першої) і повертає їх списком поруч із текстом.

    Використовується режимом --check: там важливо побачити повний список
    прогалин за один прогін, а не зупинятись на першій же помилці.
    """
    raw_text = template_path.read_text(encoding="utf-8")
    front_matter_defaults, body = split_front_matter(raw_text)
    external_data = load_data_files(data_paths)
    context = _deep_merge(_deep_merge(front_matter_defaults, external_data), cli_vars or {})

    missing: list[str] = []

    def _record(name: str | None) -> None:
        name = name or "?"
        if name not in missing:
            missing.append(name)

    class _CollectingUndefined(Undefined):
        def __str__(self) -> str:
            _record(self._undefined_name)
            return ""

        def __getattr__(self, attr: str):
            if attr.startswith("__"):
                raise AttributeError(attr)
            base = self._undefined_name or ""
            full = f"{base}.{attr}" if base else attr
            return _CollectingUndefined(hint=self._undefined_hint, obj=self._undefined_obj, name=full)

        def __iter__(self):
            _record(self._undefined_name)
            return iter([])

        def __bool__(self) -> bool:
            _record(self._undefined_name)
            return False

    env = Environment(
        undefined=_CollectingUndefined,
        autoescape=False,
        trim_blocks=True,
        lstrip_blocks=True,
    )
    template = env.from_string(body)
    rendered = template.render(**context)
    return rendered, missing




# ============================================================
# ЧАСТИНА 6: DOCX-builder (AST -> .docx)
# ============================================================


from pathlib import Path
import re

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Inches, RGBColor


MAX_IMAGE_WIDTH = Inches(6.0)
CODE_FONT = "Consolas"


class BuildWarning:
    """Збирає непопереджені під час обходу дерева проблеми (відсутні картинки,
    непідтримувані вузли) — щоб повернути їх користувачу одним списком, а не
    ховати мовчки й не валити всю збірку через один дефект."""

    def __init__(self) -> None:
        self.messages: list[str] = []

    def add(self, msg: str) -> None:
        self.messages.append(msg)


def build_docx(document: Document, output_path: Path, reference_docx: Path | None = None) -> BuildWarning:
    warnings = BuildWarning()
    doc = docx.Document(str(reference_docx)) if reference_docx else docx.Document()
    if reference_docx is None:
        _set_default_font(doc, "Times New Roman")

    for block in document.children:
        _render_block(doc, block, warnings, quote_depth=0)

    if document.footnotes:
        _render_footnotes_section(doc, document.footnotes, warnings)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return warnings


def _set_default_font(doc, font_name: str) -> None:
    """Явно прописує шрифт стилю Normal замість того, щоб лишати його
    успадкованим від docDefaults.

    Знахідка з тестування: коли шрифт рана НЕ проставлений явно (успадкований),
    LibreOffice у headless PDF-експорті на деяких сторінках великого документа
    (багато перемикань шрифтів між звичайним текстом і Consolas-кодом) іноді
    рендерить кириличний текст такого рана як tofu-квадратики — хоча в самому
    XML текст стовідсотково правильний UTF-8. Ізольований мінімальний приклад
    з тим самим patterns рендерився нормально — бага відтворюється лише в
    контексті великого документа. Явний w:rFonts на рівні стилю Normal
    (перевірено: усуває дефект) — і водночас це просто хороша практика для
    друкованих документів, де інакше шрифт залежить від того, що встановлено
    на машині користувача."""
    styles_element = doc.styles.element
    doc_defaults = styles_element.find(qn("w:docDefaults"))
    if doc_defaults is not None:
        rpr_default = doc_defaults.find(f"{qn('w:rPrDefault')}/{qn('w:rPr')}")
        if rpr_default is not None:
            existing = rpr_default.find(qn("w:rFonts"))
            if existing is not None:
                rpr_default.remove(existing)
            rFonts = OxmlElement("w:rFonts")
            for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                rFonts.set(qn(attr), font_name)
            rpr_default.insert(0, rFonts)

    normal = doc.styles["Normal"]
    normal.font.name = font_name
    rpr = normal.element.get_or_add_rPr()
    rFonts = rpr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rpr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), font_name)


# ============================================================
# Block-level рендеринг
# ============================================================

def _render_html_block_fallback(doc_or_cell, html: str, warnings: BuildWarning, quote_depth: int) -> None:
    """Сирий HTML-блок (<div>, <summary>, коментар...) сам по собі не рендериться —
    Word не виконує довільний HTML/CSS. Але раніше ми просто МОВЧКИ ВТРАЧАЛИ текст
    усередині такого блоку (напр. текст у <div>...</div>) — тепер витягуємо
    голий текст (без тегів і стилів) і друкуємо його як звичайний абзац, аби
    контент документа не зникав безслідно. Порожні/чисто-технічні блоки
    (коментарі, самотні закриваючі теги) не породжують абзаца."""
    text = _strip_html_to_text(html)
    if not text:
        return
    warnings.add(f"HTML-блок спрощено до простого тексту (втрачено форматування/стилі): {text[:60]}")
    p = doc_or_cell.add_paragraph()
    _apply_quote_indent(p, quote_depth)
    if re.match(r"^\s*<summary", html, re.IGNORECASE):
        p.add_run(text).bold = True  # <summary> — найближчий аналог: жирний "заголовок" розгортки
    else:
        p.add_run(text)


_HTML_COMMENT_RE = re.compile(r"<!--.*?-->", re.DOTALL)
_HTML_TAG_RE = re.compile(r"<[^>]+>")


def _strip_html_to_text(html: str) -> str:
    import html as html_module
    no_comments = _HTML_COMMENT_RE.sub("", html)
    no_tags = _HTML_TAG_RE.sub(" ", no_comments)
    text = html_module.unescape(no_tags)
    return " ".join(text.split())


def _render_block(doc_or_cell, node: Node, warnings: BuildWarning, quote_depth: int) -> None:
    if isinstance(node, Heading):
        p = doc_or_cell.add_paragraph(style=f"Heading {min(node.level, 9)}")
        _apply_quote_indent(p, quote_depth)
        _render_inline(p, node.children, warnings)

    elif isinstance(node, Paragraph):
        p = doc_or_cell.add_paragraph()
        _apply_quote_indent(p, quote_depth)
        _render_inline(p, node.children, warnings)

    elif isinstance(node, ThematicBreak):
        p = doc_or_cell.add_paragraph()
        _set_bottom_border(p)

    elif isinstance(node, CodeBlock):
        _render_code_block(doc_or_cell, node, quote_depth)

    elif isinstance(node, MathBlock):
        p = doc_or_cell.add_paragraph()
        _apply_quote_indent(p, quote_depth)  # pBdr/ind ПЕРЕД jc — порядок схеми OOXML
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            insert_math(p, node.latex, display=True)
        except Exception as exc:  # noqa: BLE001 — формула не має валити всю збірку
            warnings.add(f"Формула не вставлена ({exc}): {node.latex[:60]}")
            p.add_run(f"[формула: {node.latex}]").italic = True

    elif isinstance(node, HTMLBlock):
        _render_html_block_fallback(doc_or_cell, node.html, warnings, quote_depth)

    elif isinstance(node, Callout):
        _render_callout(doc_or_cell, node, warnings, quote_depth)

    elif isinstance(node, MermaidDiagram):
        _render_mermaid_diagram(doc_or_cell, node, warnings, quote_depth)

    elif isinstance(node, Blockquote):
        for child in node.children:
            _render_block(doc_or_cell, child, warnings, quote_depth + 1)

    elif isinstance(node, List):
        num_id = new_list_numid(doc_or_cell, node.ordered)
        _render_list(doc_or_cell, node, warnings, depth=0, quote_depth=quote_depth, num_id=num_id, ordered=node.ordered)

    elif isinstance(node, Table):
        _render_table(doc_or_cell, node, warnings)

    elif isinstance(node, DefinitionList):
        _render_deflist(doc_or_cell, node, warnings, quote_depth)

    elif isinstance(node, Image):
        _render_image_block(doc_or_cell, node, warnings, quote_depth)

    else:
        warnings.add(f"Невідомий block-вузол пропущено: {type(node).__name__}")


# CT_PPrBase (схема OOXML) вимагає СУВОРИЙ порядок дочірніх елементів pPr:
# ...numPr, pBdr, shd, tabs, ..., spacing, ind, contextualSpacing, ..., jc, ...
# Тому будь-яке ручне додавання w:pBdr/w:shd має відбутись ДО виклику
# python-docx властивостей paragraph_format.* (вони додають spacing/ind/jc).
# Порушення порядку -> невалідний docx, який Word може відмовитись відкрити.

def _prepare_pPr_extras(paragraph, border: bool = False, shade: str | None = None) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    if border:
        pBdr = OxmlElement("w:pBdr")
        left = OxmlElement("w:left")
        left.set(qn("w:val"), "single")
        left.set(qn("w:sz"), "12")
        left.set(qn("w:space"), "4")
        left.set(qn("w:color"), "AAAAAA")
        pBdr.append(left)
        pPr.append(pBdr)  # pBdr — перед shd
    if shade:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), shade)
        pPr.append(shd)  # shd — після pBdr, перед spacing/ind


def _apply_quote_indent(paragraph, quote_depth: int, shade: str | None = None) -> None:
    if quote_depth <= 0:
        if shade:
            _prepare_pPr_extras(paragraph, border=False, shade=shade)
        return
    _prepare_pPr_extras(paragraph, border=True, shade=shade)
    paragraph.paragraph_format.left_indent = Inches(0.3 * quote_depth)


def _set_bottom_border(paragraph) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "888888")
    pBdr.append(bottom)
    paragraph._p.get_or_add_pPr().append(pBdr)


CALLOUT_META = {
    "NOTE": {
        "label": "ПРИМІТКА",
        "icon": "ⓘ",
        "border_color": "007AFF", # macOS System Blue
        "bg_color": "F0F6FF",
        "text_rgb": RGBColor(0x00, 0x7A, 0xFF),
    },
    "TIP": {
        "label": "ПОРАДА",
        "icon": "💡",
        "border_color": "34C759", # macOS System Green
        "bg_color": "F0FAF4",
        "text_rgb": RGBColor(0x24, 0x8A, 0x3D),
    },
    "IMPORTANT": {
        "label": "ВАЖЛИВО",
        "icon": "❗",
        "border_color": "AF52DE", # macOS System Purple
        "bg_color": "FAF5FF",
        "text_rgb": RGBColor(0x89, 0x36, 0xB2),
    },
    "WARNING": {
        "label": "ЗАСТЕРЕЖЕННЯ",
        "icon": "⚠️",
        "border_color": "FF9500", # macOS System Orange
        "bg_color": "FFFBF2",
        "text_rgb": RGBColor(0xD9, 0x77, 0x06),
    },
    "CAUTION": {
        "label": "НЕБЕЗПЕКА",
        "icon": "🛑",
        "border_color": "FF3B30", # macOS System Red
        "bg_color": "FFF2F2",
        "text_rgb": RGBColor(0xD7, 0x23, 0x18),
    },
}


def _set_table_full_width(table) -> None:
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "pct")
    tblW.set(qn("w:w"), "5000")


def _set_cell_margins(cell, top: int = 140, bottom: int = 140, left: int = 200, right: int = 200) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def _set_callout_borders_and_shading(cell, border_color: str, bg_color: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")

    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "24")  # 3 pt
    left.set(qn("w:space"), "0")
    left.set(qn("w:color"), border_color)
    tcBorders.append(left)

    for side in ("top", "right", "bottom"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "none")
        tcBorders.append(el)
    tcPr.append(tcBorders)

    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), bg_color)
    tcPr.append(shd)


def _set_code_borders_and_shading(cell, border_color: str = "E2E8F0", bg_color: str = "F8FAFC") -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")  # 0.5 pt
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), border_color)
        tcBorders.append(el)
    tcPr.append(tcBorders)

    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), bg_color)
    tcPr.append(shd)


def _render_callout(doc_or_cell, node: Callout, warnings: BuildWarning, quote_depth: int) -> None:
    meta = CALLOUT_META.get(node.alert_type, CALLOUT_META["NOTE"])
    if not hasattr(doc_or_cell, "add_table"):
        for child in node.children:
            _render_block(doc_or_cell, child, warnings, quote_depth + 1)
        return

    table = doc_or_cell.add_table(rows=1, cols=1)
    table.autofit = False
    table.allow_autofit = False
    _set_table_full_width(table)

    cell = table.cell(0, 0)
    _set_callout_borders_and_shading(cell, meta["border_color"], meta["bg_color"])
    _set_cell_margins(cell, top=140, bottom=140, left=200, right=200)

    p_header = cell.paragraphs[0]
    p_header.paragraph_format.space_before = Pt(2)
    p_header.paragraph_format.space_after = Pt(4)
    run_header = p_header.add_run(f"{meta['icon']}  {meta['label']}")
    run_header.bold = True
    run_header.font.name = "Arial"
    run_header.font.size = Pt(9.5)
    run_header.font.color.rgb = meta["text_rgb"]

    for child in node.children:
        _render_block(cell, child, warnings, quote_depth=0)

    if hasattr(doc_or_cell, "add_paragraph"):
        p_spacer = doc_or_cell.add_paragraph()
        p_spacer.paragraph_format.space_before = Pt(0)
        p_spacer.paragraph_format.space_after = Pt(4)


_MERMAID_CACHE_DIR = Path("/tmp/md2docx_mermaid_cache")


def _sanitize_mermaid_code(code: str) -> str:
    """Виправляє поширені синтаксичні дефекти та застарілі конструкції Mermaid,
    щоб вони успішно парсилися сучасним рушієм.
    """
    # 1. Розділити приліплені коментарі %% після напрямку графа (graph TD%% -> graph TD\n%%)
    code = re.sub(r"^([ \t]*(?:graph|flowchart)\s+[A-Za-z]+)%%", r"\1\n%%", code, flags=re.MULTILINE)
    # 2. Сумісність для застарілих gitGraph
    if "gitGraph" in code:
        code = re.sub(r"options\s*\{[^}]*\}\s*end\s*", "", code)
        if "checkout master" in code and "branch master" not in code:
            code = code.replace("checkout master", "checkout main")
    return code


def _get_or_render_mermaid(code: str, warnings: BuildWarning) -> Path | None:
    _MERMAID_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    clean_code = _sanitize_mermaid_code(code.strip())
    code_hash = hashlib.sha256(clean_code.encode("utf-8")).hexdigest()[:20]
    cache_file = _MERMAID_CACHE_DIR / f"mermaid_{code_hash}.png"
    if cache_file.exists() and cache_file.stat().st_size > 0:
        return cache_file

    # Пріоритет 1: високонадійний сервіс Kroki (HTTP POST — без обмежень на довжину URL)
    try:
        req = urllib.request.Request(
            "https://kroki.io/mermaid/png",
            data=clean_code.encode("utf-8"),
            headers={"Content-Type": "text/plain; charset=utf-8", "User-Agent": "md2docx"},
        )
        with urllib.request.urlopen(req, timeout=12) as resp:
            data = resp.read()
        if data and len(data) > 100:
            cache_file.write_bytes(data)
            return cache_file
    except Exception as exc:
        pass  # переходимо до резервного джерела

    # Пріоритет 2: резервний сервіс mermaid.ink
    try:
        state = {"code": clean_code, "mermaid": {"theme": "default"}}
        encoded = base64.b64encode(json.dumps(state).encode("utf-8")).decode("ascii")
        url = f"https://mermaid.ink/img/{encoded}"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (md2docx)"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = resp.read()
        if data and len(data) > 100:
            cache_file.write_bytes(data)
            return cache_file
    except Exception as exc:
        warnings.add(f"Помилка рендерингу Mermaid ({exc})")

    # Пріоритет 3: локальний mmdc, якщо встановлений у системі
    import shutil
    mmdc_bin = shutil.which("mmdc")
    if mmdc_bin:
        try:
            tmp_mmd = _MERMAID_CACHE_DIR / f"temp_{code_hash}.mmd"
            tmp_mmd.write_text(clean_code, encoding="utf-8")
            res = subprocess.run(
                [mmdc_bin, "-i", str(tmp_mmd), "-o", str(cache_file), "-b", "white"],
                capture_output=True, timeout=15,
            )
            if res.returncode == 0 and cache_file.exists() and cache_file.stat().st_size > 0:
                return cache_file
        except Exception:
            pass

    return None


def _render_mermaid_diagram(doc_or_cell, node: MermaidDiagram, warnings: BuildWarning, quote_depth: int) -> None:
    img_path = _get_or_render_mermaid(node.code, warnings)
    if img_path is not None and img_path.exists():
        p = doc_or_cell.add_paragraph()
        _apply_quote_indent(p, quote_depth)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run()
        width = _scaled_width(img_path)
        run.add_picture(str(img_path), width=width)
    else:
        warnings.add(f"Діаграма Mermaid не згенерована графічно — збережено у вигляді коду: {node.code[:50]}...")
        _render_code_block(doc_or_cell, CodeBlock(code=node.code, language="mermaid"), quote_depth)


def _render_code_block(doc_or_cell, node: CodeBlock, quote_depth: int) -> None:
    lines = node.code.rstrip("\n").split("\n") or [""]
    if not hasattr(doc_or_cell, "add_table") or quote_depth > 0:
        for line in lines:
            p = doc_or_cell.add_paragraph()
            _apply_quote_indent(p, quote_depth, shade="F8FAFC")
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            run = p.add_run(line if line else " ")
            run.font.name = CODE_FONT
            run.font.size = Pt(9.0)
        return

    table = doc_or_cell.add_table(rows=1, cols=1)
    table.autofit = False
    table.allow_autofit = False
    _set_table_full_width(table)

    cell = table.cell(0, 0)
    _set_code_borders_and_shading(cell, border_color="E2E8F0", bg_color="F8FAFC")
    _set_cell_margins(cell, top=120, bottom=120, left=180, right=180)

    start_idx = 0
    if node.language:
        p_lang = cell.paragraphs[0]
        p_lang.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_lang.paragraph_format.space_before = Pt(0)
        p_lang.paragraph_format.space_after = Pt(2)
        r_lang = p_lang.add_run(node.language.upper())
        r_lang.font.name = "Arial"
        r_lang.font.size = Pt(8)
        r_lang.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)
        start_idx = 0
    else:
        p_first = cell.paragraphs[0]
        p_first.paragraph_format.space_before = Pt(0)
        p_first.paragraph_format.space_after = Pt(0)
        p_first.paragraph_format.line_spacing = 1.0
        r_first = p_first.add_run(lines[0] if lines[0] else " ")
        r_first.font.name = CODE_FONT
        r_first.font.size = Pt(9.0)
        start_idx = 1

    for line in lines[start_idx:]:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(line if line else " ")
        run.font.name = CODE_FONT
        run.font.size = Pt(9.0)

    if hasattr(doc_or_cell, "add_paragraph"):
        p_sp = doc_or_cell.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(0)
        p_sp.paragraph_format.space_after = Pt(4)


def _render_list(
    doc_or_cell, node: List, warnings: BuildWarning, depth: int, quote_depth: int, num_id: int, ordered: bool
) -> None:
    for item in node.items:
        first_block = True
        for child in item.children:
            if isinstance(child, List):
                if child.ordered != ordered:
                    # Зміна типу списку (нумерований <-> маркований) усередині вкладеності —
                    # наш multilevel abstractNum прив'язаний до ОДНОГО типу на весь numId,
                    # тому для іншого типу потрібен окремий numId (свіжий restart).
                    # depth все одно передаємо далі — щоб відступ (ilvl) відображав
                    # реальну вкладеність, а не скидався на рівень кореня.
                    child_num_id = new_list_numid(doc_or_cell, child.ordered)
                    _render_list(doc_or_cell, child, warnings, depth + 1, quote_depth, child_num_id, child.ordered)
                else:
                    _render_list(doc_or_cell, child, warnings, depth + 1, quote_depth, num_id, ordered)
                continue
            if isinstance(child, Paragraph) and first_block:
                p = doc_or_cell.add_paragraph()
                _apply_quote_indent(p, quote_depth)
                if item.checked is not None:
                    # Task-list: чекбокс замінює маркер списку повністю (не дублюємо
                    # bullet+"[x]" одночасно) — просто той самий відступ, без numPr.
                    p.paragraph_format.left_indent = Inches(0.25 * (depth + 1))
                    marker = "[x] " if item.checked else "[ ] "
                    run = p.add_run(marker)
                    run.font.name = CODE_FONT
                else:
                    apply_numbering(p, num_id, depth)
                _render_inline(p, child.children, warnings)
                first_block = False
            else:
                # вкладені блоки, що не є списком (таблиця, код, картинка...) в пункті списку
                _render_block(doc_or_cell, child, warnings, quote_depth)


def _render_table(doc_or_cell, node: Table, warnings: BuildWarning) -> None:
    if not node.rows:
        return
    n_cols = max(len(r.cells) for r in node.rows)
    table = doc_or_cell.add_table(rows=0, cols=n_cols)
    table.style = "Table Grid"
    table.autofit = True
    table.allow_autofit = True

    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }

    for row in node.rows:
        row_cells = table.add_row().cells
        for i, cell in enumerate(row.cells):
            if i >= n_cols:
                continue
            docx_cell = row_cells[i]
            p = docx_cell.paragraphs[0]
            if cell.align in align_map:
                p.alignment = align_map[cell.align]
            _render_inline(p, cell.children, warnings)
            if row.is_header:
                for run in p.runs:
                    run.bold = True
                _shade_cell(docx_cell, "E7E7E7")


def _shade_cell(cell, hex_color: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _render_deflist(doc_or_cell, node: DefinitionList, warnings: BuildWarning, quote_depth: int) -> None:
    for item in node.items:
        p_term = doc_or_cell.add_paragraph()
        _apply_quote_indent(p_term, quote_depth)
        _render_inline(p_term, item.term, warnings)
        for run in p_term.runs:
            run.bold = True
        for definition in item.definitions:
            p_def = doc_or_cell.add_paragraph()
            p_def.paragraph_format.left_indent = Inches(0.3 * (quote_depth + 1))
            _render_inline(p_def, definition, warnings)


def _render_image_block(doc_or_cell, node: Image, warnings: BuildWarning, quote_depth: int) -> None:
    p = doc_or_cell.add_paragraph()
    _apply_quote_indent(p, quote_depth)  # pBdr/ind ПЕРЕД jc — інакше порушення схеми OOXML
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    local_path = _get_local_image_path(node.src, node.resolved_path, warnings)
    if local_path is None:
        p.add_run(f"[зображення недоступне: {node.src}]").italic = True
        return
    if not local_path.exists():
        warnings.add(f"Картинку не знайдено на диску: {local_path}")
        caption = f" — {node.alt}" if node.alt else ""
        p.add_run(f"[зображення не знайдено: {node.src}{caption}]").italic = True
        return
    run = p.add_run()
    run.add_picture(str(local_path), width=_scaled_width(local_path))


_EXTERNAL_IMAGE_URL_RE = re.compile(r"^https?://", re.IGNORECASE)
_image_download_cache: dict[str, Path | None] = {}


def _get_local_image_path(src: str, resolved_path: Path | None, warnings: BuildWarning) -> Path | None:
    """Повертає локальний шлях до картинки для вбудовування — резолвлений
    парсером (локальний файл) або, для http(s)-URL, реально завантажений
    і закешований на диск (раніше зовнішні картинки просто пропускались)."""
    if resolved_path is not None:
        return resolved_path
    if _EXTERNAL_IMAGE_URL_RE.match(src):
        if src not in _image_download_cache:
            _image_download_cache[src] = _download_image(src, warnings)
        return _image_download_cache[src]
    warnings.add(f"Зображення з нелокальним/непідтримуваним джерелом пропущено: {src}")
    return None


def _download_image(url: str, warnings: BuildWarning) -> Path | None:
    import tempfile
    import urllib.request

    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (oms-convert)"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = resp.read()
            content_type = resp.headers.get("Content-Type", "")
        ext = _guess_image_extension(url, content_type)
        tmp = tempfile.NamedTemporaryFile(suffix=ext, delete=False)
        tmp.write(data)
        tmp.close()
        return Path(tmp.name)
    except Exception as exc:  # noqa: BLE001 — недоступна мережева картинка не має валити всю збірку
        warnings.add(f"Не вдалося завантажити зовнішнє зображення {url}: {exc}")
        return None


def _guess_image_extension(url: str, content_type: str) -> str:
    path_part = url.lower().split("?")[0]
    for ext in (".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"):
        if path_part.endswith(ext):
            return ext
    mapping = {
        "image/png": ".png", "image/jpeg": ".jpg", "image/gif": ".gif",
        "image/webp": ".webp", "image/bmp": ".bmp",
    }
    return mapping.get(content_type.split(";")[0].strip().lower(), ".png")


def _scaled_width(path: Path):
    try:
        from PIL import Image as PILImage
        with PILImage.open(path) as im:
            w_px, _ = im.size
        if Inches(w_px / 96) > MAX_IMAGE_WIDTH:
            return MAX_IMAGE_WIDTH
        return None  # природний розмір — не масштабуємо дрібні картинки
    except Exception:
        return MAX_IMAGE_WIDTH


def _render_footnotes_section(doc, footnotes: dict[str, FootnoteDefinition], warnings: BuildWarning) -> None:
    doc.add_paragraph(style="Heading 2").add_run("Примітки")
    for identifier, fn in footnotes.items():
        p = doc.add_paragraph()
        p.add_run(f"{identifier}. ").bold = True
        for i, child in enumerate(fn.children):
            if isinstance(child, Paragraph):
                if i > 0:
                    p = doc.add_paragraph()
                _render_inline(p, child.children, warnings)
            else:
                _render_block(doc, child, warnings, quote_depth=0)


# ============================================================
# Inline-level рендеринг
# ============================================================

_HTML_FMT_TAGS = {
    "sub": "subscript",
    "sup": "superscript",
    "u": "underline",
}


def _render_inline(paragraph, nodes: list[Node], warnings: BuildWarning, _fmt: dict | None = None) -> None:
    """Обходить inline-вузли з lookahead: <sub>/<sup>/<u> приходять з парсера як
    ДВА окремі RawInlineHTML вузли (відкриваючий + закриваючий) з текстом-сиблінгом
    між ними, а не як один вкладений вузол — тому їх треба зіставляти тут, а не
    в per-node диспетчері _render_inline_node."""
    fmt = _fmt or {}
    i = 0
    n = len(nodes)
    while i < n:
        node = nodes[i]
        if isinstance(node, RawInlineHTML):
            tag_name, is_closing = _parse_html_tag(node.html)
            if tag_name in _HTML_FMT_TAGS and not is_closing:
                close_idx = _find_matching_close(nodes, i + 1, tag_name)
                if close_idx is not None:
                    inner_fmt = {**fmt, _HTML_FMT_TAGS[tag_name]: True}
                    _render_inline(paragraph, nodes[i + 1:close_idx], warnings, inner_fmt)
                    i = close_idx + 1
                    continue
            if tag_name == "br":
                paragraph.add_run().add_break()
                i += 1
                continue
            i += 1  # непарний або незнайомий тег — мовчки пропускаємо
            continue
        _render_inline_node(paragraph, node, warnings, fmt)
        i += 1


def _parse_html_tag(html: str) -> tuple[str, bool]:
    m = re.match(r"</?\s*([a-zA-Z0-9]+)", html)
    if not m:
        return "", False
    return m.group(1).lower(), html.lstrip().startswith("</")


def _find_matching_close(nodes: list[Node], start: int, tag_name: str) -> int | None:
    depth = 1
    for j in range(start, len(nodes)):
        node = nodes[j]
        if isinstance(node, RawInlineHTML):
            name, is_closing = _parse_html_tag(node.html)
            if name == tag_name:
                if is_closing:
                    depth -= 1
                    if depth == 0:
                        return j
                else:
                    depth += 1
    return None


def _render_inline_node(paragraph, node: Node, warnings: BuildWarning, fmt: dict) -> None:
    if isinstance(node, Text):
        run = paragraph.add_run(node.value)
        _apply_fmt(run, fmt)

    elif isinstance(node, Bold):
        _render_inline(paragraph, node.children, warnings, {**fmt, "bold": True})

    elif isinstance(node, Italic):
        _render_inline(paragraph, node.children, warnings, {**fmt, "italic": True})

    elif isinstance(node, Strikethrough):
        _render_inline(paragraph, node.children, warnings, {**fmt, "strike": True})

    elif isinstance(node, InlineCode):
        run = paragraph.add_run(node.value)
        run.font.name = CODE_FONT
        _shade_run(run, "F0F0F0")
        _apply_fmt(run, fmt)

    elif isinstance(node, LineBreak):
        paragraph.add_run().add_break()

    elif isinstance(node, InlineMath):
        try:
            insert_math(paragraph, node.latex, display=False)
        except Exception as exc:  # noqa: BLE001
            warnings.add(f"Inline-формула не вставлена ({exc}): {node.latex[:60]}")
            paragraph.add_run(f"[{node.latex}]").italic = True

    elif isinstance(node, Link):
        text = "".join(_flatten_text(node.children))
        _add_hyperlink(paragraph, node.url, text or node.url)

    elif isinstance(node, AutoLink):
        _add_hyperlink(paragraph, node.url, node.url)

    elif isinstance(node, InlineImage):
        local_path = _get_local_image_path(node.src, node.resolved_path, warnings)
        if local_path and local_path.exists():
            run = paragraph.add_run()
            run.add_picture(str(local_path), width=Inches(1.5))
        else:
            paragraph.add_run(f"[{node.alt or node.src}]").italic = True

    elif isinstance(node, FootnoteRef):
        run = paragraph.add_run(node.identifier)
        run.font.superscript = True

    elif isinstance(node, RawInlineHTML):
        pass  # парні sub/sup/u/br оброблені в _render_inline(); одинокі теги сюди не доходять

    else:
        warnings.add(f"Невідомий inline-вузол пропущено: {type(node).__name__}")


def _flatten_text(nodes: list[Node]) -> list[str]:
    out: list[str] = []
    for n in nodes:
        if isinstance(n, Text):
            out.append(n.value)
        elif hasattr(n, "children"):
            out.extend(_flatten_text(n.children))
    return out


def _apply_fmt(run, fmt: dict) -> None:
    if fmt.get("bold"):
        run.bold = True
    if fmt.get("italic"):
        run.italic = True
    if fmt.get("strike"):
        run.font.strike = True
    if fmt.get("underline"):
        run.underline = True
    if fmt.get("subscript"):
        run.font.subscript = True
    if fmt.get("superscript"):
        run.font.superscript = True


def _shade_run(run, hex_color: str) -> None:
    rPr = run._r.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    rPr.append(shd)


def _add_hyperlink(paragraph, url: str, text: str):
    part = paragraph.part
    r_id = part.relate_to(
        url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rPr.append(color)
    rPr.append(underline)
    run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)
    return hyperlink


# ============================================================
# ЧАСТИНА 7: Уніфікований CLI — fill -> parse -> build, або --check
# ============================================================

def _set_nested(data: dict, dotted_key: str, value: str) -> None:
    """Встановлює значення за крапковим шляхом: 'client.name' -> {'client': {'name': value}}.
    Використовується для --var з командного рядка."""
    parts = dotted_key.split(".")
    node = data
    for part in parts[:-1]:
        node = node.setdefault(part, {})
    node[parts[-1]] = value


def _parse_var_args(pairs: list[str]) -> dict:
    """Перетворює список 'key=value' (з можливими крапками в key) на вкладений dict."""
    result: dict = {}
    for pair in pairs:
        if "=" not in pair:
            raise ValueError(f"--var очікує формат key=value, отримано: {pair!r}")
        key, value = pair.split("=", 1)
        _set_nested(result, key.strip(), value)
    return result


def convert(
    template_path: Path,
    data_paths: list[Path],
    cli_vars: dict,
    output_path: Path,
    keep_md_path: Path | None = None,
) -> list[str]:
    """Повний прогін: MD-шаблон (+дані) -> заповнений MD -> AST -> DOCX.
    Повертає список попереджень (нефатальні проблеми конвертації)."""
    raw_text = template_path.read_text(encoding="utf-8")
    front_matter_defaults, body_text = split_front_matter(raw_text)
    external_data = load_data_files(data_paths)
    context = _deep_merge(_deep_merge(front_matter_defaults, external_data), cli_vars)

    env = Environment(
        undefined=StrictUndefined,
        autoescape=False,
        trim_blocks=True,
        lstrip_blocks=True,
    )
    template = env.from_string(body_text)
    try:
        filled_text = template.render(**context)
    except UndefinedError as exc:
        raise MissingVariableError(
            f"Шаблон '{template_path.name}' посилається на змінну, якої немає в даних: {exc}"
        ) from exc

    if keep_md_path is not None:
        keep_md_path.parent.mkdir(parents=True, exist_ok=True)
        keep_md_path.write_text(filled_text, encoding="utf-8")

    document = parse_markdown(filled_text, source_path=template_path)
    warnings = build_docx(document, output_path)
    return warnings.messages


# ---- Діагностика (--check): аналіз без запису .docx ----

def _walk_ast(node) -> "list":
    """Плаский обхід усіх вузлів дерева (для підрахунку статистики)."""
    out = [node]
    for attr in ("children", "items", "rows", "cells", "definitions", "term"):
        val = getattr(node, attr, None)
        if isinstance(val, list):
            for x in val:
                out.extend(_walk_ast(x))
    return out


def _collect_images(nodes: list) -> tuple[list, list, list]:
    """Повертає (знайдені_локальні, відсутні_локальні, зовнішні_url) серед Image/InlineImage."""
    found, missing_local, external = [], [], []
    for n in nodes:
        if isinstance(n, (Image, InlineImage)):
            if n.resolved_path is None:
                external.append(n)
            elif n.resolved_path.exists():
                found.append(n)
            else:
                missing_local.append(n)
    return found, missing_local, external


def check_document(
    template_path: Path,
    data_paths: list[Path],
    cli_vars: dict,
) -> str:
    """Аналізує документ БЕЗ запису .docx і повертає текстовий звіт:
    відсутні змінні, відсутні картинки, HTML-контент що спроститься,
    загальна статистика структури. Формули НЕ проганяються через Pandoc
    у цьому режимі (щоб перевірка була швидкою навіть на великих файлах) —
    їхня валідність підтверджується вже під час реальної конвертації."""
    filled_text, missing_vars = render_template_collecting(template_path, data_paths, cli_vars)
    document = parse_markdown(filled_text, source_path=template_path)
    nodes = _walk_ast(document)

    from collections import Counter
    counts = Counter(type(n).__name__ for n in nodes)
    found_imgs, missing_imgs, external_imgs = _collect_images(nodes)
    html_blocks = [n for n in nodes if isinstance(n, HTMLBlock)]
    n_math = counts.get("InlineMath", 0) + counts.get("MathBlock", 0)

    lines: list[str] = []
    lines.append(f"=== Перевірка документа: {template_path} ===")
    lines.append("")

    lines.append("Змінні шаблону:")
    if missing_vars:
        lines.append(f"  ⚠ Відсутні значення для {len(missing_vars)} змінної(их):")
        for name in missing_vars:
            lines.append(f"      - {name}")
    else:
        lines.append("  ✓ Усі використані змінні мають значення (або їх немає в шаблоні)")
    lines.append("")

    lines.append("Структура документа:")
    lines.append(f"  Заголовків:        {counts.get('Heading', 0)}")
    lines.append(f"  Абзаців:           {counts.get('Paragraph', 0)}")
    lines.append(f"  Списків:           {counts.get('List', 0)} (пунктів: {counts.get('ListItem', 0)})")
    lines.append(f"  Таблиць:           {counts.get('Table', 0)}")
    lines.append(f"  Діаграм Mermaid:   {counts.get('MermaidDiagram', 0)}")
    lines.append(f"  Callout-блоків:    {counts.get('Callout', 0)}")
    lines.append(f"  Формул:            {n_math} ({counts.get('MathBlock', 0)} блочних, {counts.get('InlineMath', 0)} inline)")
    lines.append(f"  Блоків коду:       {counts.get('CodeBlock', 0)}")
    lines.append(f"  Цитат:             {counts.get('Blockquote', 0)}")
    lines.append(f"  Виносок:           {len(document.footnotes)}")
    lines.append(f"  Гіперпосилань:     {counts.get('Link', 0) + counts.get('AutoLink', 0)}")
    lines.append("")

    total_imgs = len(found_imgs) + len(missing_imgs) + len(external_imgs)
    lines.append(f"Зображення ({total_imgs} знайдено в тексті):")
    lines.append(f"  ✓ Локальні, знайдені на диску:      {len(found_imgs)}")
    if missing_imgs:
        lines.append(f"  ✗ Локальні, НЕ знайдені на диску:   {len(missing_imgs)}")
        for n in missing_imgs[:10]:
            lines.append(f"      - {n.src}" + (f"  ({n.alt[:60]})" if n.alt else ""))
        if len(missing_imgs) > 10:
            lines.append(f"      ... і ще {len(missing_imgs) - 10}")
    if external_imgs:
        lines.append(f"  🌐 Зовнішні URL (завантажаться під час конвертації, не перевірено зараз): {len(external_imgs)}")
        for n in external_imgs[:5]:
            lines.append(f"      - {n.src}")
        if len(external_imgs) > 5:
            lines.append(f"      ... і ще {len(external_imgs) - 5}")
    lines.append("")

    if html_blocks:
        non_empty = [b for b in html_blocks if _strip_html_to_text(b.html)]
        lines.append(f"HTML-контент, що спроститься до простого тексту (втрата стилів): {len(non_empty)}")
        for b in non_empty[:5]:
            preview = _strip_html_to_text(b.html)[:70]
            lines.append(f"      - {preview}")
        if len(non_empty) > 5:
            lines.append(f"      ... і ще {len(non_empty) - 5}")
        lines.append("")

    recs = []
    if missing_vars:
        recs.append("Додайте значення для відсутніх змінних через --data або --var.")
    if missing_imgs:
        recs.append(f"Додайте {len(missing_imgs)} файл(ів) зображень поруч зі шляхами, вказаними в документі.")
    if external_imgs:
        recs.append("Перевірте, що зовнішні URL картинок доступні з машини, де запускається конвертація.")
    if non_empty if html_blocks else False:
        recs.append("Сирий HTML (підписи, div-обгортки) втратить форматування — це очікувано, Word не виконує CSS.")
    if not recs:
        recs.append("Проблем не знайдено — документ готовий до конвертації.")

    lines.append("Рекомендації:")
    for r in recs:
        lines.append(f"  - {r}")

    return "\n".join(lines)


def main() -> None:
    parser_cli = argparse.ArgumentParser(
        prog="md2docx.py",
        description=(
            "Конвертує Markdown у DOCX із підстановкою змінних (Jinja2), формулами "
            "(LaTeX через Pandoc OMML), реальним вбудовуванням картинок (локальних і "
            "з http/https), таблицями, списками та рекурсивними цитатами."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Приклади:\n"
            "  # Просто конвертувати MD в DOCX (без змінних)\n"
            "  python3 md2docx.py document.md -o document.docx\n"
            "\n"
            "  # З підстановкою змінних із YAML-файлу (можна кілька, пізніші перекривають раніші)\n"
            "  python3 md2docx.py template.md -o out.docx --data client.yaml --data station.yaml\n"
            "\n"
            "  # З окремими значеннями напряму з командного рядка (крапкова нотація для вкладених полів)\n"
            "  python3 md2docx.py template.md -o out.docx --var client.name=\"ТОВ Ромашка\" --var act.number=17\n"
            "\n"
            "  # Комбінація — --var має найвищий пріоритет, перекриває --data і front matter\n"
            "  python3 md2docx.py template.md -o out.docx --data client.yaml --var act.number=42\n"
            "\n"
            "  # Зберегти заповнений .md для перегляду перед конвертацією\n"
            "  python3 md2docx.py template.md -o out.docx --data client.yaml --keep-md filled.md\n"
            "\n"
            "  # ПЕРЕВІРКА документа без конвертації в docx: звіт про відсутні змінні,\n"
            "  # відсутні картинки, HTML-контент що спроститься, статистику структури\n"
            "  python3 md2docx.py document.md --check\n"
            "  python3 md2docx.py template.md --check --data client.yaml\n"
            "\n"
            "Пріоритет джерел змінних (пізніше перекриває раніше):\n"
            "  front matter шаблону  →  --data (у порядку вказання)  →  --var\n"
            "\n"
            "Синтаксис шаблону — звичайний Jinja2 усередині markdown:\n"
            "  {{ client.name }}, {% for x in items %}...{% endfor %}, {% if cond %}...{% endif %}\n"
        ),
    )
    parser_cli.add_argument("template", type=Path, help="вхідний .md файл (шаблон або звичайний markdown)")
    parser_cli.add_argument(
        "-o", "--output", type=Path, default=None,
        help="шлях до вихідного .docx файлу (обов'язковий, якщо не вказано --check)",
    )
    parser_cli.add_argument(
        "--data", type=Path, action="append", default=[],
        help="YAML data-файл зі змінними (можна вказати кілька разів; пізніші перекривають раніші)",
    )
    parser_cli.add_argument(
        "--var", action="append", default=[], metavar="key=value",
        help="окреме значення змінної, крапкова нотація для вкладених полів "
             "(напр. --var client.name=Foo); має найвищий пріоритет, перекриває --data і front matter",
    )
    parser_cli.add_argument(
        "--keep-md", type=Path, default=None,
        help="зберегти заповнений (проміжний) .md файл для перегляду перед конвертацією",
    )
    parser_cli.add_argument(
        "--check", action="store_true",
        help="ТІЛЬКИ перевірити документ і надрукувати звіт (відсутні змінні, відсутні "
             "картинки, HTML-контент що спроститься, статистика структури) — без "
             "запису .docx. -o/--output не потрібен у цьому режимі. Формули не "
             "перевіряються через Pandoc у цьому режимі (для швидкості на великих файлах).",
    )
    parser_cli.add_argument(
        "-v", "--verbose", action="store_true",
        help="детальний покроковий вивід виконання (verbose)",
    )
    if len(sys.argv) <= 1:
        parser_cli.print_help()
        sys.exit(0)

    args = parser_cli.parse_args()

    if not args.check and args.output is None:
        parser_cli.error("-o/--output є обов'язковим, якщо не вказано --check")

    try:
        cli_vars = _parse_var_args(args.var)
        if args.check:
            report = check_document(args.template, args.data, cli_vars)
            print(report)
            return
        msgs = convert(args.template, args.data, cli_vars, args.output, args.keep_md)
    except (FileNotFoundError, ValueError, MissingVariableError) as exc:
        print(f"Помилка: {exc}", file=sys.stderr)
        sys.exit(1)

    print(f"Готово: {args.output}")
    if args.keep_md:
        print(f"Заповнений .md збережено: {args.keep_md}")
    if msgs:
        print(f"\nПопередження ({len(msgs)}) — конвертація завершена, але зверніть увагу:")
        for m in msgs:
            print(f"  - {m}")


if __name__ == "__main__":
    main()

