"""
Парсер .docx документов (КонсультантПлюс и др.) в JSON для RAG / нейронки.

Структура одного фрагмента в JSON:
{
    "fragment_index": int,
    "type":          str,    # "text" | "table" | "footnote" | "title"
    "text":          str,    # для table -- markdown-таблица (один фрагмент);
                             # для остального -- обычный текст (≤ MAX_TEXT_CHARS=512)
    "header":        str,    # путь по иерархии заголовков
    "page_number":   int,    # lastRenderedPageBreak / page break + эвристика по символам
    "point":         str,    # маркер пункта в начале параграфа ("1.", "2.1.", "1)", ...)
    "note":          str,    # текст примечания / сноски, привязанной к фрагменту
    "url":           str,    # гиперссылки из фрагмента (через "\n")
    "fz_number":     str,    # "от 21.07.1997 N 116-ФЗ"
    "fz_text":       str,    # "О промышленной безопасности опасных производственных объектов"
    "images":        [str],  # имена PNG, лежащих рядом с этим JSON
    "is_indexable":  bool    # False -- служебный/юридический блок (преамбула,
                             # "Признать утратившими силу", подписи, шапки приложения,
                             # "Зарегистрировано в Минюсте"...). Такие фрагменты
                             # НЕ нужно класть в индекс RAG -- они засоряют top-k.
}

Использование:
    python parse_docs.py                       # DOKI -> OUT
    python parse_docs.py DOKI OUT
    python parse_docs.py DOKI/file.docx OUT
"""

from __future__ import annotations

import json
import os
import re
import shutil
import sys
from dataclasses import dataclass, field, asdict
from typing import Iterable

from docx import Document
from docx.oxml.ns import qn

# --- Параметры ---------------------------------------------------------------

MAX_TEXT_CHARS = 512      # ограничение на длину text для обычных параграфов
SOFT_CHARS = 480          # стараемся резать чуть раньше, чтобы не рвать слова
CHARS_PER_PAGE = 2800     # эвристика для page_number, если в docx нет page-break'ов
GARBAGE_MIN_CHARS = 30    # text короче этого без point/url/note/images считается мусором

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


# --- Регулярки для бизнес-полей ----------------------------------------------

# Полная конструкция: "Федерального закона от 21.07.1997 N 116-ФЗ "О ...""
# Берём номер + (опционально) идущее следом название в кавычках.
FZ_RE = re.compile(
    r"(?:Федеральн\w+\s+закон\w*\s+)?"            # "Федерального закона" / "Федеральным законом" — опционально
    r"(от\s+\d{1,2}\.\d{1,2}\.\d{2,4}\s+N\s*[\d\-\.\/]+\s*-?\s*ФЗ)"  # сам номер
    r'(?:\s*"([^"]{3,300})")?',                   # в кавычках текст ФЗ
    re.IGNORECASE,
)

# Маркер пункта в начале параграфа: "1.", "2.3.", "10.1.2.", "1)", "а)", "II."
POINT_RE = re.compile(
    r"^\s*("
    r"\d+(?:\.\d+)*\.?"                # 1   1.   1.2   1.2.3.
    r"|\d+\)"                          # 1)
    r"|[а-яa-z]\)"                     # а) б) a)
    r"|[IVX]+\."                       # II.  III.
    r")\s+",
    re.IGNORECASE,
)

# Сноска в тексте: "<1>", "<23>" — ссылка на примечание в конце документа
FOOTNOTE_REF_RE = re.compile(r"<(\d{1,3})>")

# Маркер примечания: "КонсультантПлюс: примечание."
NOTE_LINE_RE = re.compile(r"КонсультантПлюс\s*:\s*примечан", re.IGNORECASE)


# --- Структуры данных --------------------------------------------------------

@dataclass
class Block:
    """Логический блок: текст + метаданные ДО нарезки на чанки."""
    type: str = "text"            # "text" | "table" | "footnote" | "title"
    text: str = ""
    header: str = ""
    page_number: int = 1
    point: str = ""
    note: str = ""
    urls: list[str] = field(default_factory=list)
    fz_number: str = ""
    fz_text: str = ""
    images: list[str] = field(default_factory=list)
    is_indexable: bool = True     # False -- служебные блоки (преамбула,
                                  # «Признать утратившими силу», подписи, ...).
                                  # Их не нужно класть в индекс RAG.
    # служебное
    is_note_block: bool = False
    is_footnote: bool = False
    footnote_num: str = ""
    no_split: bool = False        # не резать на 512-символьные чанки
                                  # (используется после continuation-merge)


# --- Утилиты -----------------------------------------------------------------

def iter_body_elements(doc) -> Iterable:
    """Идём по дочерним элементам body в естественном порядке (параграфы и таблицы)."""
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag in ("p", "tbl"):
            yield tag, child


def paragraph_text(p_elem) -> str:
    """Текст параграфа со склейкой всех <w:t>."""
    parts = []
    for t in p_elem.iter(qn("w:t")):
        if t.text:
            parts.append(t.text)
    return "".join(parts).strip()


def paragraph_style(p_elem) -> str:
    style_elem = p_elem.find(qn("w:pPr") + "/" + qn("w:pStyle"))
    if style_elem is not None:
        return style_elem.get(qn("w:val")) or ""
    return ""


def paragraph_hyperlinks(p_elem, rels) -> list[str]:
    """Все таргеты гиперссылок параграфа (http/https + consultantplus://)."""
    urls = []
    for hl in p_elem.iter(qn("w:hyperlink")):
        rid = hl.get(qn("r:id"))
        if not rid:
            continue
        rel = rels.get(rid)
        if rel is None:
            continue
        target = rel.target_ref or ""
        if target and target.startswith(("http://", "https://", "consultantplus://")):
            urls.append(target)
    return urls


def paragraph_page_breaks(p_elem) -> int:
    """Сколько page-break внутри этого параграфа (последний рендер Word + жёсткие break)."""
    cnt = 0
    for br in p_elem.iter(qn("w:lastRenderedPageBreak")):
        cnt += 1
    for br in p_elem.iter(qn("w:br")):
        if br.get(qn("w:type")) == "page":
            cnt += 1
    return cnt


def paragraph_image_ids(p_elem) -> list[str]:
    """rId всех картинок (a:blip @r:embed) внутри параграфа/ячейки."""
    ids = []
    for blip in p_elem.iter(qn("a:blip") if False else f"{{{A_NS}}}blip"):
        rid = blip.get(f"{{{R_NS}}}embed")
        if rid:
            ids.append(rid)
    return ids


def cell_image_ids(cell_elem) -> list[str]:
    ids = []
    for blip in cell_elem.iter(f"{{{A_NS}}}blip"):
        rid = blip.get(f"{{{R_NS}}}embed")
        if rid:
            ids.append(rid)
    return ids


def is_title_style(style: str) -> bool:
    """Заголовки: ConsPlusTitle, Heading1..., Title и т.п.
    Намеренно НЕ считаем заголовком ConsPlusTitlePage (это шапка КонсультантПлюс).
    """
    if not style:
        return False
    s = style.lower()
    if s == "consplustitlepage":
        return False
    return ("title" in s) or s.startswith("heading") or s in {"заголовок", "subtitle"}


def is_footer_separator(text: str) -> bool:
    """Длинная строка из дефисов — отделяет блок сносок в конце документа."""
    t = text.strip()
    return len(t) >= 5 and set(t) <= set("-—–")


def _cell_text(cell_elem) -> str:
    """Текст ячейки с сохранением переноса между параграфами."""
    parts = []
    for p in cell_elem.iter(qn("w:p")):
        txt = "".join(t.text or "" for t in p.iter(qn("w:t"))).strip()
        if txt:
            parts.append(txt)
    return " ".join(parts).strip()


def parse_table_matrix(tbl_elem) -> tuple[list[list[str]], list[str]]:
    """
    Парсим таблицу с учётом <w:gridSpan> (colspan) и <w:vMerge> (rowspan).
    Возвращаем (матрица текстов одинаковой ширины, [rId картинок]).

    Алгоритм vMerge: ячейка с w:vMerge без атрибута val (или val="continue")
    продолжает вертикальное объединение со стороны restart-ячейки сверху.
    Мы дублируем содержимое restart-ячейки во все продолжающие позиции
    (что-то более удобное для нейронки, чем "пусто").
    """
    image_ids: list[str] = []
    raw_rows: list[list[dict]] = []  # [{text, colspan, vmerge: None|'restart'|'continue'}, ...]

    for tr in tbl_elem.iterchildren(qn("w:tr")):
        row_cells = []
        for tc in tr.iterchildren(qn("w:tc")):
            text = _cell_text(tc)
            image_ids.extend(cell_image_ids(tc))

            # tcPr/gridSpan
            colspan = 1
            vmerge = None
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is not None:
                gs = tcPr.find(qn("w:gridSpan"))
                if gs is not None:
                    try:
                        colspan = max(1, int(gs.get(qn("w:val")) or "1"))
                    except ValueError:
                        colspan = 1
                vm = tcPr.find(qn("w:vMerge"))
                if vm is not None:
                    vmerge = vm.get(qn("w:val")) or "continue"
            row_cells.append({"text": text, "colspan": colspan, "vmerge": vmerge})
        if row_cells:
            raw_rows.append(row_cells)

    if not raw_rows:
        return [], image_ids

    # Разворачиваем colspan: каждая ячейка занимает `colspan` позиций.
    # Текст пишем в первую позицию, остальные позиции — копия (для kv-таблиц
    # это иногда удобнее, чем "").
    expanded: list[list[str]] = []
    expanded_vmerge: list[list[str | None]] = []
    for row in raw_rows:
        row_text: list[str] = []
        row_vm: list[str | None] = []
        for c in row:
            row_text.append(c["text"])
            row_vm.append(c["vmerge"])
            for _ in range(c["colspan"] - 1):
                row_text.append(c["text"])
                row_vm.append(c["vmerge"])
        expanded.append(row_text)
        expanded_vmerge.append(row_vm)

    # Выровняем ширину строк (на случай разной grid-ширины)
    width = max(len(r) for r in expanded)
    for r, vm in zip(expanded, expanded_vmerge):
        while len(r) < width:
            r.append("")
            vm.append(None)

    # vMerge: ячейка с vmerge in {'continue', None-but-text-empty-after-restart-above}
    # должна унаследовать текст у ближайшего restart-предка по той же колонке.
    for col in range(width):
        last_restart_text = ""
        for ri in range(len(expanded)):
            vm = expanded_vmerge[ri][col]
            txt = expanded[ri][col]
            if vm == "restart":
                last_restart_text = txt
            elif vm == "continue":
                expanded[ri][col] = last_restart_text
            else:
                last_restart_text = txt if txt else last_restart_text

    return expanded, image_ids


def _classify_table(matrix: list[list[str]]) -> str:
    """
    Эвристическая классификация таблицы.

    - 'single': все строки имеют только одну осмысленную колонку → текстовая обёртка.
    - 'kv':     2 эффективные колонки (label | value) или у всех строк col0 -
                короткий лейбл (<= 80 символов), col1..N — длинное значение.
    - 'data':   3+ колонок, первая строка похожа на header (короткие тексты,
                нет очень длинных) → header + data-строки.
    - 'mixed':  по умолчанию — рендерим row-by-row без header.
    """
    if not matrix:
        return "single"
    width = len(matrix[0])
    # эффективная ширина — сколько различных колонок реально несут текст
    nonempty_widths = []
    for r in matrix:
        nonempty_widths.append(sum(1 for c in r if c.strip()))
    max_nonempty = max(nonempty_widths) if nonempty_widths else 0

    if width <= 1 or max_nonempty <= 1:
        return "single"

    if width == 2:
        return "kv"

    # 3+ колонок: kv если первая колонка везде короткая и одинаковая роль
    # (часто бывает: label | val_short | val_long).
    if width >= 3:
        first_col_short = all(
            len(r[0]) <= 80 for r in matrix if r[0].strip()
        )
        # data-таблица — если первая строка по виду header (все короткие)
        first_row = matrix[0]
        looks_like_header = (
            all(len(c) <= 60 for c in first_row)
            and sum(1 for c in first_row if c.strip()) >= 2
        )
        if looks_like_header and not first_col_short:
            return "data"
        if first_col_short:
            return "kv"
        return "data" if looks_like_header else "mixed"

    return "mixed"


# Декоративные/служебные ячейки таблиц, которые НЕ несут информации:
# "(наименование)", "(код)", "(код ОКЗ <1>)", и т.п. — обычно идут отдельной
# строкой-пояснением под основной строкой с данными.
_DECOR_CELL_RE = re.compile(r"^\s*\([^()]{1,80}\)\s*$")
_TRIVIAL_CELL_VALUES = {"", "-", "—", "–", "x", "X", "*"}


def _is_decorative_cell(s: str) -> bool:
    """Ячейка — пустая, '-', '—', или текст целиком в скобках типа '(наименование)'."""
    t = s.strip()
    if t in _TRIVIAL_CELL_VALUES:
        return True
    if _DECOR_CELL_RE.match(t):
        return True
    return False


def _is_decorative_row(row: list[str]) -> bool:
    """Вся строка состоит только из декоративных/пустых ячеек."""
    return all(_is_decorative_cell(c) for c in row)


def _clean_table_matrix(matrix: list[list[str]]) -> list[list[str]]:
    """
    Очистка матрицы таблицы перед рендером:
    - выкидываем строки-пояснения (вся строка из '(...)' / пустые / '—');
    - выкидываем подряд идущие дубли (артефакты vMerge);
    - выкидываем колонки, в которых все ячейки декоративные;
    - в каждой строке декоративные ячейки превращаем в '' (но строку
      сохраняем, если есть хотя бы одна осмысленная).
    """
    if not matrix:
        return matrix

    # 1) Подряд идущие дубли строк
    dedup: list[list[str]] = []
    for row in matrix:
        if dedup and dedup[-1] == row:
            continue
        dedup.append(row)
    matrix = dedup

    # 2) Выкидываем строки-пояснения
    matrix = [row for row in matrix if not _is_decorative_row(row)]
    if not matrix:
        return []

    # 3) Колонки, в которых все ячейки декоративные → удалить
    width = max(len(r) for r in matrix)
    keep_cols: list[int] = []
    for c in range(width):
        col_values = [(r[c] if c < len(r) else "") for r in matrix]
        if any(not _is_decorative_cell(v) for v in col_values):
            keep_cols.append(c)
    if not keep_cols:
        return []
    cleaned = [[(row[c] if c < len(row) else "") for c in keep_cols] for row in matrix]

    # 4) Декоративные ячейки заменяем на ""
    for row in cleaned:
        for i, c in enumerate(row):
            if _is_decorative_cell(c):
                row[i] = ""

    # 5) повторный dedup после очистки
    dedup2: list[list[str]] = []
    for row in cleaned:
        if dedup2 and dedup2[-1] == row:
            continue
        if not any(c.strip() for c in row):
            continue
        dedup2.append(row)

    # 6) Если все строки идентичны — это шапка без данных, оставим одну.
    if not dedup2:
        return []
    unique_rows = {tuple(r) for r in dedup2}
    if len(unique_rows) == 1 and len(dedup2) > 1:
        return [dedup2[0]]
    return dedup2


def _md_escape_cell(s: str) -> str:
    """Экранирование ячейки в markdown-таблице."""
    return s.replace("|", "\\|").replace("\n", " ").strip()


def _compact_row(row: list[str]) -> list[str]:
    """Сворачиваем подряд идущие дубли (последствие colspan-разворота)."""
    out: list[str] = []
    for c in row:
        if out and out[-1] == c:
            continue
        out.append(c)
    return out


def _matrix_to_markdown(matrix: list[list[str]], kind: str) -> str:
    """
    Превращает матрицу в одну markdown-таблицу.
    Для kind='kv' использует шапку 'Параметр / Значение'.
    Для kind='data' первая строка матрицы — это header.
    Для kind='single' просто текст.
    """
    if not matrix:
        return ""

    if kind == "single":
        return "\n".join(
            next((c for c in r if c.strip()), "") for r in matrix
        ).strip()

    if kind == "kv":
        # Сжимаем дубли в каждой строке (colspan/vMerge артефакты)
        compact_rows: list[list[str]] = []
        for r in matrix:
            r2 = _compact_row([c.strip() for c in r])
            # пустые строки выкидываем
            if not any(r2):
                continue
            compact_rows.append(r2)
        if not compact_rows:
            return ""
        # Ширина = max ширина среди строк
        width = max(len(r) for r in compact_rows)
        # Заголовок: "Параметр | Значение | Значение | ..."
        if width == 2:
            header_cells = ["Параметр", "Значение"]
        else:
            header_cells = ["Параметр"] + ["Значение"] * (width - 1)
        lines = [
            "| " + " | ".join(header_cells) + " |",
            "|" + "|".join(["---"] * width) + "|",
        ]
        for r in compact_rows:
            padded = r + [""] * (width - len(r))
            lines.append("| " + " | ".join(_md_escape_cell(c) for c in padded) + " |")
        return "\n".join(lines)

    if kind == "data":
        # Первая строка — header
        header = [c.strip() for c in matrix[0]]
        # дедуп подряд идущих дублей в header (gridSpan-артефакт)
        compact_header = _compact_row(header)
        width = max(len(compact_header), max((len(r) for r in matrix[1:]), default=0))
        # выровняем header
        header_out = compact_header + ["—"] * (width - len(compact_header))
        lines = [
            "| " + " | ".join(_md_escape_cell(c) if c else "—" for c in header_out) + " |",
            "|" + "|".join(["---"] * width) + "|",
        ]
        for row in matrix[1:]:
            cells = _compact_row([c.strip() for c in row])
            if not any(cells):
                continue
            padded = cells + [""] * (width - len(cells))
            lines.append("| " + " | ".join(_md_escape_cell(c) for c in padded) + " |")
        return "\n".join(lines)

    # mixed / fallback
    compact_rows = []
    for r in matrix:
        r2 = _compact_row([c.strip() for c in r])
        if any(r2):
            compact_rows.append(r2)
    if not compact_rows:
        return ""
    width = max(len(r) for r in compact_rows)
    lines = []
    # первая строка как header
    h = compact_rows[0] + [""] * (width - len(compact_rows[0]))
    lines.append("| " + " | ".join(_md_escape_cell(c) if c else "—" for c in h) + " |")
    lines.append("|" + "|".join(["---"] * width) + "|")
    for r in compact_rows[1:]:
        padded = r + [""] * (width - len(r))
        lines.append("| " + " | ".join(_md_escape_cell(c) for c in padded) + " |")
    return "\n".join(lines)


def render_table_markdown(tbl_elem) -> tuple[str, str, list[str]]:
    """
    Главный рендер таблицы.
    Возвращает: (kind, text, [rId картинок])
      kind = 'table'  -> text это markdown-таблица
      kind = 'text'   -> text это plain text (single-таблица или после очистки осталась 1 строка)
      kind = ''       -> таблица пуста после очистки
    Правило: 1 таблица = 1 fragment, не режется на чанки.
    """
    matrix, image_ids = parse_table_matrix(tbl_elem)
    if not matrix:
        return "", "", image_ids

    matrix = _clean_table_matrix(matrix)
    if not matrix:
        return "", "", image_ids

    kind = _classify_table(matrix)
    md = _matrix_to_markdown(matrix, kind)
    if not md.strip():
        return "", "", image_ids

    # single — это не таблица, отдадим как plain text-блок
    if kind == "single":
        return "text", md, image_ids

    # маленькая таблица без markdown-маркера — тоже текст
    if "|" not in md:
        return "text", md, image_ids

    return "table", md, image_ids


def extract_fz(text: str) -> tuple[str, str]:
    """Первое попавшееся упоминание ФЗ -> (fz_number, fz_text)."""
    m = FZ_RE.search(text)
    if not m:
        return "", ""
    return (m.group(1) or "").strip(), (m.group(2) or "").strip()


def _normalize_for_dedup(s: str) -> str:
    """Нормализация для определения дублей: убираем регистр и лишние пробелы."""
    return re.sub(r"\s+", " ", s).strip().lower()


def _dedup_blocks(blocks: list["Block"]) -> list["Block"]:
    """
    Дедуп блоков. Стратегия зависит от типа:

    - type='table': дедуп ПО ТЕКСТУ глобально (без учёта header). Если
      одна и та же markdown-таблица встречается в нескольких секциях
      (типичный случай — пустой шаблон-форма) — оставляем только первое
      появление.
    - остальные типы: ключ (header, normalize(text)). Совпадение текста
      в РАЗНЫХ разделах — не дубль (контекст разный, информация полезна).

    Полностью пустые блоки без картинок выкидываем.
    """
    seen_text_table: set[str] = set()
    seen_other: set[tuple[str, str]] = set()
    out: list[Block] = []
    for b in blocks:
        if not b.text.strip() and not b.images:
            continue

        if b.type == "table":
            norm = _normalize_for_dedup(b.text)
            if norm in seen_text_table:
                # картинки переносим в первый такой же блок
                for prev in reversed(out):
                    if prev.type == "table" and _normalize_for_dedup(prev.text) == norm:
                        for img in b.images:
                            if img not in prev.images:
                                prev.images.append(img)
                        break
                continue
            seen_text_table.add(norm)
            out.append(b)
            continue

        key = (b.header.strip(), _normalize_for_dedup(b.text))
        if key in seen_other:
            for prev in reversed(out):
                if (prev.header.strip(), _normalize_for_dedup(prev.text)) == key:
                    for img in b.images:
                        if img not in prev.images:
                            prev.images.append(img)
                    break
            continue
        seen_other.add(key)
        out.append(b)
    return out


def _strip_header_from_text(blocks: list["Block"]) -> list["Block"]:
    """
    Убираем дублирование header в text:
    - если text начинается с того же, что в последнем сегменте header,
      обрезаем повтор;
    - применяется только к type='text' (table = markdown, его не трогаем;
      title = шапка документа, тоже не трогаем).
    """
    for b in blocks:
        if b.type != "text" or not b.text or not b.header:
            continue
        last_seg = b.header.split(" > ")[-1].strip()
        if not last_seg:
            continue
        normalized_text = b.text.lstrip()
        if normalized_text.lower().startswith(last_seg.lower()):
            # отрезаем повтор + возможный разделитель (": ", "\n", " — ")
            cut = len(last_seg)
            rest = normalized_text[cut:].lstrip(" :—–-\n\r\t")
            if rest:
                b.text = rest
    return blocks


def _is_garbage_block(b: "Block") -> bool:
    """
    Мусорный мини-фрагмент: type=text/footnote, текст < GARBAGE_MIN_CHARS
    символов, нет point/url/note/images, нет упоминания ФЗ.
    """
    if b.type in ("table", "title"):
        return False
    if b.images or b.point or b.urls or b.note or b.fz_number:
        return False
    if b.is_footnote:
        return False
    return len(b.text.strip()) < GARBAGE_MIN_CHARS


def _cleanup_garbage(blocks: list["Block"]) -> list["Block"]:
    """
    Удаляем мусорные мини-фрагменты:
    - пытаемся слить со следующим того же header'а (если не table),
    - иначе со ВЫШЕстоящим того же header'а,
    - иначе удаляем.
    """
    out: list[Block] = []
    i = 0
    while i < len(blocks):
        b = blocks[i]
        if _is_garbage_block(b):
            # 1) попробовать слить со следующим того же header
            nxt = blocks[i + 1] if i + 1 < len(blocks) else None
            if nxt and nxt.type == "text" and nxt.header == b.header \
                    and len(b.text) + 1 + len(nxt.text) <= MAX_TEXT_CHARS:
                nxt.text = (b.text + "\n" + nxt.text).strip()
                i += 1
                continue
            # 2) попробовать слить с последним вышестоящим того же header
            if out and out[-1].type == "text" and out[-1].header == b.header \
                    and len(out[-1].text) + 1 + len(b.text) <= MAX_TEXT_CHARS:
                out[-1].text = (out[-1].text + "\n" + b.text).strip()
                i += 1
                continue
            # 3) иначе — просто удаляем
            i += 1
            continue
        out.append(b)
        i += 1
    return out


# --- Классификация: индексировать ли фрагмент в RAG ---------------------------

# Префиксы, после которых ВЕСЬ фрагмент — служебный/юридический мусор.
# Сравниваем по нижнему регистру, .lstrip() от текста.
_NON_INDEXABLE_PREFIXES = (
    "зарегистрировано в минюсте",
    "зарегистрировано министерством юстиции",
    "список изменяющих документов",
    "признать утратившими силу",
    "утратил силу",
    "утратили силу",
    "утвердить правила",
    "утвердить порядок",
    "утвердить положение",
    "утвердить прилагаем",
    "приложение к приказу",
    "приложение n",
    "приложение №",
)

# Текст «приказываю:» в конце — это вступительная формула приказа.
_PREAMBLE_PHRASE_RE = re.compile(r"приказыва(?:ю|ем)\s*:?\s*$", re.IGNORECASE)

# «(в ред. Приказа Минтруда России от ...)» — редакционная пометка, целиком.
_EDIT_NOTE_RE = re.compile(r"^\s*\(\s*в\s+ред(?:\.|акции)\s+приказ", re.IGNORECASE)

# ФИО подписанта вида «А.О.КОТЯКОВ», «И.И.ИВАНОВ-ПЕТРОВ» и т.п.
_SIGNATORY_INITIALS_RE = re.compile(r"^[А-ЯЁ]\.[А-ЯЁ]\.[А-ЯЁ\-]{2,}$")

# Должности подписантов и адресные реквизиты приложения.
_SIGNATORY_TITLES = {
    "министр", "заместитель министра", "директор", "руководитель",
    "и.о. министра", "и.о. директора",
}

# Шапки приложения «к приказу Министерства труда / Российской Федерации /
# от 14 июля 2021 г. N 467н».
_APPENDIX_HEADER_PATTERNS = (
    re.compile(r"^\s*к\s+приказу\s+министерства", re.IGNORECASE),
    re.compile(r"^\s*российской\s+федерации\s*$", re.IGNORECASE),
    re.compile(r"^\s*и\s+социальной\s+защиты\s*$", re.IGNORECASE),
    re.compile(r"^\s*от\s+\d{1,2}\s+\w+\s+\d{4}\s+г\.\s+N\s*\d+", re.IGNORECASE),
)


def _is_non_indexable_text(b: "Block") -> bool:
    """Проверяет один блок на признак служебного / юридического мусора."""
    t = b.text.strip()
    if not t:
        return True
    low = t.lower()

    # 1) известные служебные префиксы
    for pref in _NON_INDEXABLE_PREFIXES:
        if low.startswith(pref):
            return True

    # 2) вступительная формула «В соответствии с ... приказываю:»
    if low.startswith("в соответствии с") and _PREAMBLE_PHRASE_RE.search(low):
        return True
    if _PREAMBLE_PHRASE_RE.search(low) and len(t) < 600:
        return True

    # 3) редакционная пометка целиком, без другого текста
    if _EDIT_NOTE_RE.match(t) and t.rstrip().endswith(")") and len(t) < 300:
        return True

    # 4) подпись: должность ("Министр") или ФИО ("А.О.КОТЯКОВ")
    if len(t) <= 60:
        if low in _SIGNATORY_TITLES:
            return True
        if _SIGNATORY_INITIALS_RE.match(t):
            return True

    # 5) короткие реквизиты приложения
    if len(t) <= 120:
        for rx in _APPENDIX_HEADER_PATTERNS:
            if rx.match(t):
                return True

    return False


def _mark_indexable(blocks: list["Block"]) -> list["Block"]:
    """
    Помечаем служебные блоки `is_indexable=False`.

    Особый случай: после блока «Признать утратившими силу:» идёт перечень
    приказов, каждый — отдельный фрагмент. Помечаем весь перечень целиком,
    пока не встретим блок с `point` (новый пункт) или таблицу.
    """
    in_revoked_list = False
    for b in blocks:
        # type=title -- это преамбула документа, в индекс не идёт
        if b.type == "title":
            b.is_indexable = False
            continue

        # сноски (footnote) -- индексируем, обычно полезная инфа

        t = b.text.strip()
        low = t.lower()

        # вход в перечень утративших силу
        if low.startswith("признать утратившими силу"):
            b.is_indexable = False
            in_revoked_list = True
            continue

        # пока внутри перечня "Признать утратившими силу:"
        if in_revoked_list:
            # выход из перечня -- новый пункт (нумерованный) или таблица
            if b.point or b.type in ("table", "footnote"):
                in_revoked_list = False
            else:
                # типичные элементы перечня: "приказ Министерства ...",
                # "пункт 27 изменений ...", "Министерством юстиции ..."
                if (low.startswith("приказ ")
                        or low.startswith("пункт ")
                        or low.startswith("министерством юстиции")
                        or low.startswith("n ")):
                    b.is_indexable = False
                    continue
                # иначе считаем что вышли из перечня
                in_revoked_list = False

        # общие критерии "это служебный блок"
        if _is_non_indexable_text(b):
            b.is_indexable = False

    return blocks


# Признак "продолжения" текста в следующем параграфе.
_OPEN_TAIL_RE = re.compile(r"[(\[«„,;:\-]\s*$")
_LOWER_START_RE = re.compile(r"^\s*[a-zа-яё]")
_CLOSING_BRACKET_START_RE = re.compile(r"^\s*[)\]»“]")
_END_OF_SENTENCE = set(".!?…»\")")


def _has_unclosed_bracket(s: str) -> bool:
    """В строке открытых скобок/кавычек больше чем закрытых -> фраза не закончена.

    Проверяем:
    - парные скобки: () [] «» „"
    - симметричную двойную кавычку \": если количество нечётное, значит висит
      открытая кавычка.
    """
    pairs = [("(", ")"), ("[", "]"), ("«", "»"), ("„", "“")]
    for op, cl in pairs:
        if s.count(op) > s.count(cl):
            return True
    # двойная прямая кавычка — нечётное количество = висит открытая
    if s.count('"') % 2 == 1:
        return True
    return False


def _is_continuation(a_text: str, b_text: str) -> bool:
    """Эвристика: параграф B — продолжение разорванной фразы из A."""
    a = a_text.rstrip()
    b = b_text.lstrip()
    if not a or not b:
        return False
    # 1) A заканчивается на открывающую скобку/запятую/тире
    if _OPEN_TAIL_RE.search(a):
        return True
    # 2) B начинается с закрывающей скобки — явное продолжение скобочного выражения
    if _CLOSING_BRACKET_START_RE.match(b):
        return True
    # 3) B начинается со строчной буквы И A не на знаке конца предложения
    if a[-1] not in _END_OF_SENTENCE and _LOWER_START_RE.match(b):
        return True
    # 4) В A висит незакрытая скобка/кавычка — фраза НЕ закончена
    if a[-1] not in _END_OF_SENTENCE and _has_unclosed_bracket(a):
        return True
    return False


def _can_merge(a: "Block", b: "Block", continuation: bool = False) -> bool:
    """Можно ли слепить два соседних блока в один (semantic chunking).
    continuation=True ослабляет лимит длины — потому что лучше длинный
    цельный фрагмент, чем рваная фраза.
    """
    if a.type != b.type:
        return False
    # таблицы и сноски — атомарные единицы
    if a.type in ("table", "footnote", "title"):
        return False
    if a.header != b.header:
        return False
    if a.is_footnote or b.is_footnote:
        return False
    # маркеры пункта = атомарные единицы, не сливаем
    if b.point:
        return False
    # картинки слипать нельзя (они привязаны к конкретному блоку)
    if b.images:
        return False
    total = len(a.text) + 1 + len(b.text)
    if continuation:
        # для продолжения разорванной фразы — лимит на грабли x2,
        # потому что цельность фразы важнее формального лимита
        if total > MAX_TEXT_CHARS * 2:
            return False
    else:
        if total > MAX_TEXT_CHARS:
            return False
    return True


def _semantic_merge(blocks: list["Block"]) -> list["Block"]:
    """
    Объединяем соседние короткие блоки одного раздела в один.
    Дополнительно: если B — это продолжение разорванной фразы из A
    (висячая открытая скобка, запятая в конце A, строчная буква в начале B),
    склеиваем даже при превышении MAX_TEXT_CHARS.
    """
    if not blocks:
        return blocks
    out: list[Block] = [blocks[0]]
    for b in blocks[1:]:
        last = out[-1]
        cont = (last.type == "text" and b.type == "text"
                and last.header == b.header
                and _is_continuation(last.text, b.text))
        if _can_merge(last, b, continuation=cont):
            # для продолжения слепляем БЕЗ "\n" — это одна фраза
            sep = "" if cont else ("\n" if last.text else "")
            if cont and last.text and not last.text.endswith(" ") and not b.text.startswith((" ", ")", "]", ",", ".", ";", ":", "»", "”")):
                sep = " "
            last.text = (last.text + sep + b.text).strip()
            if b.urls:
                merged_urls = list(dict.fromkeys(last.urls + b.urls))
                last.urls = merged_urls
            if b.note:
                last.note = (last.note + "\n" + b.note).strip() if last.note else b.note
            if not last.fz_number and b.fz_number:
                last.fz_number = b.fz_number
                last.fz_text = b.fz_text
            # если это была continuation-склейка и результат вышел за лимит —
            # не позволяем splitter'у позже резать этот блок обратно
            if cont and len(last.text) > MAX_TEXT_CHARS:
                last.no_split = True
        else:
            out.append(b)
    return out


def split_into_chunks(text: str, max_chars: int = MAX_TEXT_CHARS, soft: int = SOFT_CHARS) -> list[str]:
    """Режем длинный текст по ≤ max_chars символов, стараясь по границе предложения/слова.

    Особый случай: если в тексте на момент потенциального разреза есть
    незакрытая скобка/кавычка (т.е. фраза в скобках разорвётся посередине),
    смещаем границу разреза дальше — до закрытия скобки или конца текста.
    Это критично для перечислений типа «приказ ... (зарегистрирован ... N 12345)».
    """
    text = text.strip()
    if not text:
        return []
    if len(text) <= max_chars:
        return [text]

    chunks: list[str] = []
    i = 0
    n = len(text)
    while i < n:
        end = min(i + max_chars, n)
        if end < n:
            window = text[i:end]
            # сначала по концу предложения
            chosen_pos = -1
            for sep in (". ", "! ", "? ", "; ", "\n"):
                pos = window.rfind(sep)
                if pos >= soft:
                    chosen_pos = i + pos + len(sep)
                    break
            if chosen_pos < 0:
                pos = window.rfind(" ")
                if pos >= soft:
                    chosen_pos = i + pos + 1
            if chosen_pos > 0:
                end = chosen_pos

            # Защита от разрыва внутри парных структур
            open_chars = "([«„"
            close_chars = ")]»“"
            segment = text[i:end]

            # начальный depth и состояние кавычки из УЖЕ накопленного segment
            depth = 0
            for ch in segment:
                if ch in open_chars:
                    depth += 1
                elif ch in close_chars and depth > 0:
                    depth -= 1
            quote_open = (segment.count('"') % 2 == 1)

            # двигаемся вперёд пока есть открытая структура ИЛИ
            # пока сразу за нами идёт новая открытая скобка/кавычка
            k = end
            while k < n:
                if depth == 0 and not quote_open:
                    # проверим: следующий значимый символ — открывающий?
                    p = k
                    while p < n and text[p] in " \t\n":
                        p += 1
                    if p < n and (text[p] in open_chars or text[p] == '"'):
                        # двинемся к нему и продолжим — это продолжение фразы
                        k = p
                    else:
                        break
                ch = text[k]
                if ch == '"':
                    quote_open = not quote_open
                elif ch in open_chars:
                    depth += 1
                elif ch in close_chars and depth > 0:
                    depth -= 1
                k += 1

            # после закрытия структуры — дотянем до конца текущего слова
            if k > end:
                while k < n and text[k] not in " ;\n":
                    k += 1
                if k < n and text[k] in " ;":
                    k += 1
                end = k

        chunks.append(text[i:end].strip())
        i = end
    return [c for c in chunks if c]


# --- Сборщик блоков из документа ---------------------------------------------

class DocxParser:
    def __init__(self, docx_path: str, out_dir: str):
        self.docx_path = docx_path
        self.out_dir = out_dir
        self.doc_name = os.path.splitext(os.path.basename(docx_path))[0]

        self.doc = Document(docx_path)
        self.rels = self.doc.part.rels  # rId -> Relationship

        # heading-стек, чтобы строить иерархический header
        self.header_stack: list[tuple[str, str]] = []  # [(level_tag, text)]
        self.current_page = 1

        # эвристика страниц: если в docx нет page-break'ов вообще, накапливаем
        # объём текста и каждые ~CHARS_PER_PAGE символов перелистываем страницу.
        self._chars_since_page_break = 0
        self._has_explicit_page_breaks = False

        # преамбула: первые ConsPlusTitle до начала "контентной" части
        self.preamble_titles: list[str] = []
        self._preamble_open = True   # пока не встретили L1-секцию ("I. ...")

        # сноски <N> -> текст
        self.footnotes: dict[str, str] = {}
        self.image_counter = 0
        self.saved_images: dict[str, str] = {}  # rId -> имя файла

    def _advance_page_by_chars(self, n: int) -> None:
        """Эвристика: каждые ~CHARS_PER_PAGE символов = +1 страница."""
        if self._has_explicit_page_breaks:
            return
        self._chars_since_page_break += n
        while self._chars_since_page_break >= CHARS_PER_PAGE:
            self.current_page += 1
            self._chars_since_page_break -= CHARS_PER_PAGE

    def _hard_page_break(self, n: int = 1) -> None:
        """Жёсткие page-break'и (lastRenderedPageBreak / w:br type=page)."""
        if n <= 0:
            return
        self._has_explicit_page_breaks = True
        self.current_page += n
        self._chars_since_page_break = 0

    # ----- иерархия заголовков -------------------------------------------

    @staticmethod
    def heading_level(text: str) -> str:
        """
        Грубо определяем уровень заголовка по форме нумерации.
        I. ...        -> 'L1'
        3.1. ...      -> 'L2'
        3.1.1. ...    -> 'L3'
        Иначе         -> 'L0'
        """
        t = text.strip()
        if re.match(r"^[IVX]+\.\s", t):
            return "L1"
        m = re.match(r"^(\d+(?:\.\d+)*)\.\s", t)
        if m:
            depth = m.group(1).count(".") + 2  # "3" -> L2, "3.1" -> L3, "3.1.1" -> L4
            return f"L{depth}"
        return "L0"

    def push_header(self, text: str):
        lvl = self.heading_level(text)
        # выталкиваем равные или более глубокие уровни
        order = ["L1", "L2", "L3", "L4", "L5", "L0"]
        idx = order.index(lvl) if lvl in order else len(order) - 1
        self.header_stack = [(l, t) for (l, t) in self.header_stack
                             if order.index(l) < idx]
        self.header_stack.append((lvl, text))

    def current_header(self) -> str:
        return " > ".join(t for _, t in self.header_stack)

    def _flush_preamble(self) -> "Block":
        """Создаёт title-блок с шапкой документа.
        В header кладём краткое название документа (имя файла без префикса '11. '),
        в text — содержимое преамбулы. Это исключает дублирование header в text.
        """
        titles = [t for t in self.preamble_titles if t.strip()]
        text = "\n".join(titles).strip()
        # короткое название документа = имя файла без префикса нумерации
        short_name = re.sub(r"^\d+\.\s*", "", self.doc_name).strip() or self.doc_name
        blk = Block(
            type="title",
            text=text,
            header=short_name,
            page_number=1,
        )
        self.preamble_titles = []
        return blk

    # ----- сохранение картинок -------------------------------------------

    def save_image(self, rid: str, fragment_index: int) -> str | None:
        """Сохраняет картинку по rId, возвращает имя файла (без пути) или None."""
        if rid in self.saved_images:
            return self.saved_images[rid]
        rel = self.rels.get(rid)
        if rel is None or "image" not in rel.reltype.lower():
            return None
        try:
            blob = rel.target_part.blob
        except Exception:
            return None

        # расширение
        target = rel.target_ref or ""
        ext = os.path.splitext(target)[1].lower() or ".png"
        if ext not in (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".emf", ".wmf"):
            ext = ".png"
        # принудительно .png по ТЗ (кроме явных форматов оставим как есть, иначе нейронка не откроет EMF)
        if ext in (".emf", ".wmf"):
            ext = ".png"  # как заглушка; конвертация EMF/WMF в PNG требует доп.инструмента

        self.image_counter += 1
        fname = f"schema{self.image_counter}_{self.doc_name}_{fragment_index}{ext}"
        fpath = os.path.join(self.out_dir, fname)
        try:
            with open(fpath, "wb") as f:
                f.write(blob)
        except OSError:
            return None
        self.saved_images[rid] = fname
        return fname

    # ----- основной проход ------------------------------------------------

    def parse(self) -> list[Block]:
        blocks: list[Block] = []
        pending_note: str = ""           # примечание -> прикрепится к следующему контентному блоку
        last_was_title: bool = False     # последний обработанный параграф был title?

        for tag, elem in iter_body_elements(self.doc):
            if tag == "p":
                style = paragraph_style(elem)
                text = paragraph_text(elem)

                # учёт страниц (page-break ВНУТРИ параграфа считаем ПЕРЕД самим блоком)
                pb = paragraph_page_breaks(elem)
                self._hard_page_break(pb)

                if not text:
                    last_was_title = False
                    continue

                # разделительная строка перед сносками
                if is_footer_separator(text):
                    last_was_title = False
                    continue

                # 1) Служебная строка КонсультантПлюс — пропускаем сразу
                if text.lower().startswith("документ предоставлен консультантплюс"):
                    last_was_title = False
                    continue

                # 2) Это заголовок?
                if is_title_style(style):
                    # если ещё в преамбуле — собираем заголовки в неё (МИНИСТЕРСТВО /
                    # ПРИКАЗ / от ... / ОБ УТВЕРЖДЕНИИ ... / "СПЕЦИАЛИСТ ...")
                    if self._preamble_open and self.heading_level(text) != "L1":
                        self.preamble_titles.append(text)

                    # если предыдущий блок был тоже title без пустой строки — это
                    # продолжение того же заголовка
                    if last_was_title and self.header_stack:
                        prev_lvl, prev_text = self.header_stack[-1]
                        self.header_stack[-1] = (prev_lvl, f"{prev_text} {text}".strip())
                    else:
                        self.push_header(text)

                    # как только встретили L1 ("I. ...") — преамбула закрыта,
                    # выгружаем её отдельным блоком type="title"
                    if self.heading_level(text) == "L1" and self._preamble_open:
                        self._preamble_open = False
                        if self.preamble_titles:
                            blocks.append(self._flush_preamble())
                    last_was_title = True
                    self._advance_page_by_chars(len(text))
                    continue

                # сюда дошли — это не title и не служебка
                last_was_title = False

                # 3) Это маркер примечания "КонсультантПлюс: примечание."
                if NOTE_LINE_RE.search(text):
                    pending_note = (pending_note + " " + text).strip() if pending_note else text
                    continue

                # 4) Это сноска <N> ... ?
                m_fn = re.match(r"^\s*<(\d{1,3})>\s*(.+)$", text)
                if m_fn:
                    num = m_fn.group(1)
                    body = m_fn.group(2).strip()
                    self.footnotes[num] = body
                    blk = self._make_block(elem, body)
                    blk.type = "footnote"
                    blk.is_footnote = True
                    blk.footnote_num = num
                    blk.header = (self.current_header() + " > Примечания").strip(" >")
                    blk.note = f"<{num}>"
                    blocks.append(blk)
                    self._advance_page_by_chars(len(body))
                    continue

                # если до сих пор в преамбуле, но это уже обычный текст —
                # закрываем преамбулу и выгружаем заголовки как title
                if self._preamble_open and self.preamble_titles:
                    self._preamble_open = False
                    blocks.append(self._flush_preamble())

                # 5) Обычный контентный параграф
                blk = self._make_block(elem, text)
                if pending_note:
                    blk.note = (blk.note + "\n" + pending_note).strip()
                    pending_note = ""
                blocks.append(blk)
                self._advance_page_by_chars(len(text))

            elif tag == "tbl":
                last_was_title = False
                kind, md, image_ids = render_table_markdown(elem)

                if md:
                    blk = Block(
                        type=kind,            # "table" или "text"
                        text=md,
                        header=self.current_header(),
                        page_number=self.current_page,
                    )
                    blk.fz_number, blk.fz_text = extract_fz(md)
                    refs = sorted(set(FOOTNOTE_REF_RE.findall(md)))
                    if refs:
                        blk.note = "\n".join(f"<{r}>" for r in refs)
                    if pending_note:
                        blk.note = (blk.note + "\n" + pending_note).strip()
                        pending_note = ""
                    if image_ids:
                        blk.images = image_ids
                        image_ids = []
                    blocks.append(blk)
                    self._advance_page_by_chars(len(md))

                # если остались картинки без текста — отдельный контейнер
                if image_ids:
                    blk = Block(
                        type="text",
                        text="",
                        header=self.current_header(),
                        page_number=self.current_page,
                        images=image_ids,
                    )
                    if pending_note:
                        blk.note = pending_note
                        pending_note = ""
                    blocks.append(blk)

        # если документ короткий и преамбула так и не закрылась — закроем сейчас
        if self._preamble_open and self.preamble_titles:
            self._preamble_open = False
            blocks.insert(0, self._flush_preamble())

        # title-блок переносим в начало (на случай если он создался после
        # пустых строк)
        title_header = ""
        for i, b in enumerate(blocks):
            if b.type == "title":
                if i != 0:
                    blocks.insert(0, blocks.pop(i))
                title_header = b.header
                break

        # блокам с пустым header (например, "Зарегистрировано в Минюсте..."
        # идёт до начала преамбулы) ставим header из title-блока
        if title_header:
            for b in blocks:
                if b.type != "title" and not b.header.strip():
                    b.header = title_header

        # дорезолвим ссылки на сноски: если в blk.note стоят "<N>", подставим текст
        for blk in blocks:
            if blk.is_footnote:
                continue
            text_for_refs = blk.text + " " + blk.note
            refs = sorted(set(FOOTNOTE_REF_RE.findall(text_for_refs)))
            if refs:
                resolved = []
                for r in refs:
                    if r in self.footnotes:
                        resolved.append(f"<{r}> {self.footnotes[r]}")
                if resolved:
                    blk.note = "\n".join(resolved)

        blocks = _semantic_merge(blocks)
        blocks = _strip_header_from_text(blocks)
        blocks = _dedup_blocks(blocks)
        blocks = _cleanup_garbage(blocks)
        blocks = _mark_indexable(blocks)
        return blocks

    def _make_block(self, p_elem, text: str) -> Block:
        # point — маркер в начале текста
        point = ""
        m = POINT_RE.match(text)
        clean_text = text
        if m:
            point = m.group(1).strip()
            clean_text = text[m.end():].strip()

        urls = paragraph_hyperlinks(p_elem, self.rels)
        fz_number, fz_text = extract_fz(clean_text)
        img_ids = paragraph_image_ids(p_elem)

        return Block(
            text=clean_text,
            header=self.current_header(),
            page_number=self.current_page,
            point=point,
            urls=urls,
            fz_number=fz_number,
            fz_text=fz_text,
            images=img_ids,
        )


# --- Главная функция парсинга одного файла ----------------------------------

def parse_one_docx(docx_path: str, out_dir: str) -> str:
    os.makedirs(out_dir, exist_ok=True)
    parser = DocxParser(docx_path, out_dir)
    blocks = parser.parse()

    # Block -> фрагменты.
    # Таблицы (type="table") и заглавный блок (type="title") НЕ режутся
    # на 512-символьные куски — у них свой лимит/логика.
    # Обычный текст режется на чанки ≤ MAX_TEXT_CHARS.
    fragments = []
    next_index = 0
    for blk in blocks:
        if blk.type == "table":
            chunks = [blk.text] if blk.text else [""]
        elif blk.type == "title":
            chunks = [blk.text] if blk.text else [""]
        elif blk.no_split:
            # continuation-склейка — фраза не должна разрываться обратно
            chunks = [blk.text] if blk.text else [""]
        else:
            chunks = split_into_chunks(blk.text) if blk.text else [""]
            if not chunks:
                chunks = [""]

        for i, chunk_text in enumerate(chunks):
            frag = {
                "fragment_index": next_index,
                "type": blk.type,
                "text": chunk_text,
                "header": blk.header,
                "page_number": blk.page_number,
                "point": blk.point if i == 0 else "",
                "note": blk.note if i == 0 else "",
                "url": "\n".join(blk.urls) if i == 0 else "",
                "fz_number": blk.fz_number,
                "fz_text": blk.fz_text,
                "images": [],
                "is_indexable": blk.is_indexable,
            }
            if i == 0 and blk.images:
                saved_names = []
                for rid in blk.images:
                    fname = parser.save_image(rid, next_index)
                    if fname:
                        saved_names.append(fname)
                frag["images"] = saved_names
            fragments.append(frag)
            next_index += 1

    out_json_path = os.path.join(out_dir, f"{parser.doc_name}.json")
    with open(out_json_path, "w", encoding="utf-8") as f:
        json.dump(fragments, f, ensure_ascii=False, indent=2)

    return out_json_path


# --- CLI ---------------------------------------------------------------------

def main():
    src = sys.argv[1] if len(sys.argv) > 1 else "DOKI"
    dst = sys.argv[2] if len(sys.argv) > 2 else "OUT"

    if os.path.isfile(src):
        files = [src]
    elif os.path.isdir(src):
        files = [os.path.join(src, f) for f in os.listdir(src)
                 if f.lower().endswith(".docx") and not f.startswith("~$")]
    else:
        print(f"Не найдено: {src}")
        sys.exit(1)

    if not files:
        print(f"В {src} нет .docx файлов")
        sys.exit(1)

    os.makedirs(dst, exist_ok=True)
    for fp in files:
        print(f"[+] {fp}")
        out = parse_one_docx(fp, dst)
        with open(out, "r", encoding="utf-8") as f:
            data = json.load(f)
        print(f"    -> {out}  фрагментов: {len(data)}")


if __name__ == "__main__":
    main()
