#!/usr/bin/env python3
"""
PDF -> Word (DOCX): простой конвертер с распознаванием сканов (OCR) и прогрессом.

    python pdf2word.py документ.pdf              -> документ.docx рядом с PDF, сразу открывается
    python pdf2word.py документ.pdf итог.docx
    python pdf2word.py                           -> откроется окно выбора PDF

Страницы с текстовым слоем конвертирует pdf2docx (сохраняются шрифты, таблицы,
картинки). Страницы-сканы распознаёт Tesseract OCR: текст, линии таблиц и картинки
переносятся на «чистую» PDF-страницу, которую затем так же разбирает pdf2docx —
получаются обычные редактируемые абзацы, таблицы и рисунки.
"""
import argparse
import copy
import functools
import getpass
import logging
import os
import re
import shutil
import statistics
import subprocess
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

import cv2
import numpy as np
import pymupdf
import pytesseract
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Emu, Pt
from docx.text.paragraph import Paragraph
from PIL import Image
from tqdm import tqdm

# служебные сообщения PyMuPDF и INFO-логи pdf2docx ломают прогресс-бар — прячем их
pymupdf.set_messages(pylogging=True, pylogging_level=logging.DEBUG)
from pdf2docx import Converter  # noqa: E402
from pdf2docx.text.Line import Line  # noqa: E402
logging.getLogger().setLevel(logging.ERROR)


def _make_line_with_space(self, p, _make_line=Line.make_docx):
    """pdf2docx склеивает строки абзаца без пробела («Покупатель» + «обязуется» ->
    «Покупательобязуется»), если в PDF строка не заканчивается пробелом. Добавляем его."""
    _make_line(self, p)
    if not self.line_break and p.runs:
        texts = p.runs[-1]._r.findall(qn("w:t"))
        if texts and texts[-1].text and not texts[-1].text.endswith((" ", "-", "\u00ad")):
            texts[-1].text += " "
            texts[-1].set(qn("xml:space"), "preserve")


Line.make_docx = _make_line_with_space

APP_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))  # папка программы (или распакованного exe)
BUNDLED_TESSERACT = APP_DIR / "tesseract"  # Tesseract, встроенный в PDF2Word.exe (см. build_exe.py)
OCR_FONTS = {False: APP_DIR / "assets" / "fonts" / "LiberationSerif-Regular.ttf",  # метрики как у
             True: APP_DIR / "assets" / "fonts" / "LiberationSerif-Bold.ttf"}       # Times New Roman
MIN_CHARS = 20          # меньше букв/цифр на странице -> считаем её сканом
SIZE_FACTOR = 1.08      # x_size строки от Tesseract -> кегль шрифта (подобрано на тестах)
XHTML = "{http://www.w3.org/1999/xhtml}"
# «Вес» шагов в прогресс-баре, чтобы проценты шли примерно равномерно по времени
COST_OCR, COST_ANALYZE, COST_PARSE, COST_BUILD = 25, 1, 2, 1

# маркеры списков из шрифтов Symbol/Wingdings (частная область Unicode) -> обычные символы
SYMBOL_BULLETS = {"\uf0b7": "•", "\uf0a7": "▪", "\uf0d8": "➢", "\uf076": "❖",
                  "\uf0fc": "✓", "\uf0e8": "➔", "\uf06e": "■", "\uf071": "❑"}
# свободные шрифты, которыми PDF часто заменяет шрифты Windows -> их оригиналы
FONT_NAMES = {"liberationserif": "Times New Roman", "liberationsans": "Arial", "liberationmono": "Courier New"}
# так Tesseract нередко читает маркер списка «•»
BULLET_LIKE = {"•", "·", "●", "®", "©", "*", "»", "«", "°", "¢", "e", "o", "е", "о"}

WINDOWS_TESSERACT = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
]


# ----------------------------------------------------------------------------- OCR

def setup_tesseract(lang):
    """Находит tesseract.exe и оставляет только установленные языки."""
    if (BUNDLED_TESSERACT / "tesseract.exe").exists():
        pytesseract.pytesseract.tesseract_cmd = str(BUNDLED_TESSERACT / "tesseract.exe")
        os.environ["TESSDATA_PREFIX"] = str(BUNDLED_TESSERACT / "tessdata")
    elif not shutil.which("tesseract"):
        for path in WINDOWS_TESSERACT:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                break
    try:
        installed = set(pytesseract.get_languages())
    except pytesseract.TesseractNotFoundError:
        sys.exit("Для распознавания сканов нужен Tesseract OCR, но он не найден.\n"
                 "Установите его: https://github.com/UB-Mannheim/tesseract/wiki "
                 "(при установке отметьте язык Russian).")
    langs = [code for code in lang.split("+") if code in installed]
    missing = [code for code in lang.split("+") if code not in installed]
    if missing:
        tqdm.write(f"Внимание: языки OCR не установлены и будут пропущены: {', '.join(missing)}")
    if not langs:
        sys.exit(f"Ни один из языков '{lang}' не установлен в Tesseract. Есть: {', '.join(sorted(installed))}")
    return "+".join(langs)


def has_text_layer(page):
    """True, если на странице есть видимый и читаемый текст (а не только картинка-скан)."""
    try:  # невидимый текст (type 3) — это чужой OCR-слой поверх скана, его не считаем
        chars = [chr(c[0]) for span in page.get_texttrace() if span["type"] != 3 for c in span["chars"]]
    except Exception:
        chars = list(page.get_text())
    good = sum(ch.isalnum() for ch in chars)
    return good >= MIN_CHARS and chars.count("\ufffd") < good


def hocr_value(el, key):
    """Числа из title-атрибута hOCR: 'bbox 10 20 30 40; x_size 47' -> [10, 20, 30, 40]."""
    m = re.search(rf"{key} ([^;]+)", el.get("title", ""))
    return [float(v) for v in m.group(1).split()] if m else None


def rotate(img, angle, border=255):
    """Поворачивает изображение на angle градусов против часовой стрелки."""
    h, w = img.shape[:2]
    matrix = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1)
    return cv2.warpAffine(img, matrix, (w, h), flags=cv2.INTER_LINEAR, borderValue=(border,) * 3)


def skew_angle(gray):
    """Наклон скана: при верном угле строки текста дают самые резкие «пики» в сумме по строкам."""
    small = cv2.resize(gray, None, fx=0.25, fy=0.25, interpolation=cv2.INTER_AREA)
    ink = (small < 160).astype(np.float32)

    def sharpness(angle):
        return float(np.var(rotate(ink, angle, border=0).sum(axis=1))), -abs(angle)  # при равенстве — 0°

    coarse = max(np.arange(-5, 5.01, 0.5), key=sharpness)
    return float(max(np.arange(coarse - 0.5, coarse + 0.51, 0.05), key=sharpness))


def stroke_width(ink, box):
    """Толщина штриха букв: медиана длин горизонтальных отрезков «чернил» в рамке строки."""
    x0, y0, x1, y1 = (int(v) for v in box)
    edges = np.diff(np.pad(ink[y0:y1, x0:x1], ((0, 0), (1, 1))).astype(np.int8), axis=1)
    runs = np.nonzero(edges == -1)[1] - np.nonzero(edges == 1)[1]
    return float(np.median(runs)) if runs.size else 0.0


def find_rules(ink, dpi):
    """Линии таблиц и подчёркивания: длинные тонкие горизонтальные и вертикальные отрезки (в пикселях)."""
    rules, mask = [], np.zeros_like(ink)
    for kernel in ((dpi // 2, 1), (1, dpi * 3 // 10)):
        found = cv2.morphologyEx(ink, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, kernel))
        for contour in cv2.findContours(found, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)[0]:
            x, y, w, h = cv2.boundingRect(contour)
            if min(w, h) > dpi / 20:
                continue  # толстое — это заливка или картинка, а не линия
            rules.append((x, y + h / 2, x + w, y + h / 2, h) if w > h else (x + w / 2, y, x + w / 2, y + h, w))
            cv2.rectangle(mask, (x, y), (x + w - 1, y + h - 1), 255, -1)
    return rules, cv2.dilate(mask, np.ones((5, 5), np.uint8))


def join_hyphenated(lines):
    """«сло-» в конце строки + «во» в начале следующей -> «слово», иначе перенос мешает правке текста."""
    for a, b in zip(lines, lines[1:]):
        if not a["words"] or not b["words"]:
            continue
        last, first = a["words"][-1], b["words"][0]
        if len(last["text"]) > 2 and last["text"].endswith("-") and last["text"][-2].isalpha() \
                and first["text"][:1].islower():
            last["text"] = last["text"][:-1] + first["text"]
            b["words"].pop(0)
    return [line for line in lines if line["words"]]


def words_box(words):
    """Общая рамка нескольких слов."""
    return [min(w["box"][0] for w in words), min(w["box"][1] for w in words),
            max(w["box"][2] for w in words), max(w["box"][3] for w in words)]


def recover_missed(clean, ink, rules_mask, lines, lang, dpi):
    """Дочитывает то, что Tesseract пропустил: одиночные цифры в ячейках таблиц, номера страниц и т.п.
    Все такие кусочки собираются полосками на одну картинку и распознаются за один запуск."""
    if not lines:
        return []
    x_size = int(statistics.median(line["x_size"] for line in lines))
    left = ink.copy()
    left[rules_mask > 0] = 0
    near = x_size // 2 + 4  # кусочки вплотную к словам — это «хвосты» уже распознанных букв (№, Р...)
    for line in lines:
        for word in line["words"]:
            x0, y0, x1, y1 = (int(v) for v in word["box"])
            left[max(y0 - 4, 0):y1 + 4, max(x0 - near, 0):x1 + near] = 0
    merged = cv2.dilate(left, np.ones((3, max(x_size * 3 // 5, 1)), np.uint8))  # буквы слова — в один кусок
    crops, pad = [], x_size // 3
    for x, y, w, h, _ in cv2.connectedComponentsWithStats(merged)[2][1:]:
        if 0.4 * x_size <= h <= 1.8 * x_size and w <= dpi * 2 and len(crops) < 50:
            x0, y0 = max(x - pad, 0), max(y - pad, 0)
            crops.append((x0, y0, clean[y0:y + h + pad, x0:x + w + pad]))
    if not crops:
        return []

    canvas = np.full((sum(c.shape[0] + x_size for *_, c in crops) + x_size,
                      max(c.shape[1] for *_, c in crops) + 2 * x_size), 255, np.uint8)
    bands, top = [], x_size
    for x0, y0, crop in crops:
        canvas[top:top + crop.shape[0], x_size:x_size + crop.shape[1]] = crop
        bands.append((top, top + crop.shape[0], x0 - x_size, y0 - top))
        top += crop.shape[0] + x_size
    data = pytesseract.image_to_data(Image.fromarray(canvas), lang=lang, config="--oem 1 --psm 6",
                                     output_type=pytesseract.Output.DICT)
    found = {}
    for i, text in enumerate(data["text"]):
        text, conf = text.strip(), float(data["conf"][i])
        if not text or conf < 70 or not any(ch.isalnum() for ch in text):
            continue
        middle = data["top"][i] + data["height"][i] / 2
        for band, (y_top, y_bottom, dx, dy) in enumerate(bands):
            if y_top <= middle <= y_bottom:
                box = [data["left"][i] + dx, data["top"][i] + dy,
                       data["left"][i] + data["width"][i] + dx, data["top"][i] + data["height"][i] + dy]
                found.setdefault(band, []).append({"box": box, "text": text, "conf": conf})
    dark = clean < 128
    return [{"words": sorted(words, key=lambda w: w["box"][0]), "base": max(w["box"][3] for w in words),
             "x_size": x_size, "stroke": stroke_width(dark, words_box(words)) / x_size}
            for words in found.values()]


def recognize(page, lang, dpi):
    """Распознаёт страницу-скан: строки слов с кеглем и жирностью, линии таблиц и картинки."""
    pix = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csRGB, alpha=False)
    color = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, 3)
    gray = cv2.medianBlur(cv2.cvtColor(color, cv2.COLOR_RGB2GRAY), 3)  # убирает «пыль» скана
    try:  # автоповорот страниц, отсканированных боком или вверх ногами
        turn = pytesseract.image_to_osd(Image.fromarray(gray), output_type=pytesseract.Output.DICT)["rotate"]
    except pytesseract.TesseractError:
        turn = 0  # слишком мало текста для определения ориентации
    if turn:
        color, gray = (np.ascontiguousarray(np.rot90(img, -turn // 90)) for img in (color, gray))
    angle = skew_angle(gray)
    if abs(angle) > 0.05:  # выравниваем наклон: строки и линии таблиц становятся строго горизонтальными
        color, gray = rotate(color, angle), rotate(gray, angle)

    ink = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 51, 20)
    edge = int(dpi * 0.15)  # тени и края листа у границы скана — не содержимое
    ink[:edge], ink[-edge:], ink[:, :edge], ink[:, -edge:] = 0, 0, 0, 0
    rules, rules_mask = find_rules(ink, dpi)
    clean = gray.copy()
    clean[rules_mask > 0] = 255  # без линий Tesseract намного лучше читает текст в ячейках таблиц

    hocr = pytesseract.image_to_pdf_or_hocr(Image.fromarray(clean), lang=lang, config="--oem 1 --psm 3",
                                            extension="hocr")
    dark = clean < 128
    lines = []
    for par in ET.fromstring(hocr).iter(f"{XHTML}p"):
        par_lines = []
        for line in par:
            words = []
            for w in line:
                text = "".join(w.itertext()).strip()
                conf = (hocr_value(w, "x_wconf") or [100])[0]
                if text and (conf >= 30 or any(ch.isalnum() for ch in text)):  # иначе это грязь скана
                    words.append({"box": hocr_value(w, "bbox"), "text": text, "conf": conf})
            if not words:
                continue
            box = hocr_value(line, "bbox")
            x_size = max((hocr_value(line, "x_size") or [box[3] - box[1]])[0], 1)
            if len(words) > 1 and words[0]["text"] in BULLET_LIKE \
                    and words[1]["box"][0] - words[0]["box"][2] > 0.8 * x_size:
                words[0]["text"] = "•"
            par_lines.append({"words": words, "base": box[3] + (hocr_value(line, "baseline") or [0, 0])[1],
                              "x_size": x_size, "stroke": stroke_width(dark, box) / x_size})
        lines += join_hyphenated(par_lines)
    lines += recover_missed(clean, ink, rules_mask, lines, lang, dpi)

    # жирный шрифт: штрихи заметно толще, чем у основного текста страницы
    normal = statistics.median(line["stroke"] for line in lines) if lines else 0
    for line in lines:
        line["bold"] = line["stroke"] > normal * 1.3

    # картинки (фото, графики, печати, подписи): всё заметно темнее бумаги, кроме слов и линий
    paper = float(np.percentile(gray, 95))
    rest = np.where(gray < paper - 45, 255, 0).astype(np.uint8)
    rest[:edge], rest[-edge:], rest[:, :edge], rest[:, -edge:] = 0, 0, 0, 0
    rest[rules_mask > 0] = 0
    for line in lines:
        for word in line["words"]:
            x0, y0, x1, y1 = (int(v) for v in word["box"])
            rest[max(y0 - 4, 0):y1 + 4, max(x0 - 4, 0):x1 + 4] = 0
    blobs = cv2.dilate(rest, np.ones((25, 25), np.uint8))
    boxes = [[x, y, x + w, y + h] for x, y, w, h, _ in cv2.connectedComponentsWithStats(blobs)[2][1:]
             if min(w, h) >= dpi // 10]
    gap, merged = dpi // 2, True
    while merged:  # части одной картинки (столбики графика, печать и подпись рядом) — в одну
        merged = False
        for a in boxes:
            b = next((b for b in boxes if b is not a and a[0] - gap <= b[2] and b[0] - gap <= a[2]
                      and a[1] - gap <= b[3] and b[1] - gap <= a[3]), None)
            if b:
                a[:] = min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3])
                boxes.remove(b)
                merged = True
                break
    figures = []
    for x0, y0, x1, y1 in boxes:
        if x1 - x0 >= dpi * 0.4 and y1 - y0 >= dpi * 0.3 and rest[y0:y1, x0:x1].mean() > 255 * 0.03:
            crop = cv2.cvtColor(color[y0:y1, x0:x1], cv2.COLOR_RGB2BGR)
            figures.append(((x0, y0, x1, y1), cv2.imencode(".png", crop)[1].tobytes()))

    def in_figure(word):  # «слова», которые Tesseract нашёл внутри картинки, обычно мусор
        cx, cy = (word["box"][0] + word["box"][2]) / 2, (word["box"][1] + word["box"][3]) / 2
        return any(f[0] <= cx <= f[2] and f[1] <= cy <= f[3] for f, _ in figures)

    for line in lines:
        line["words"] = [w for w in line["words"] if w["conf"] >= 70 or not in_figure(w)]
    lines = [line for line in lines if line["words"]]
    rules = [r for r in rules if not in_figure({"box": r[:4]})]  # края столбиков графика — не таблица

    # кегль: по ширине слов (шрифт метрически как Times New Roman), проверяя по высоте строки
    k = 72 / dpi
    for line in lines:
        by_height = line["x_size"] * k * SIZE_FACTOR
        natural = sum(ocr_font(line["bold"]).text_length(w["text"], 1) for w in line["words"])
        by_width = sum(w["box"][2] - w["box"][0] for w in line["words"]) * k / natural if natural else 0
        line["size"] = by_width if 0.75 < by_width / by_height < 1.35 else by_height
    # кегль основного текста немного «гуляет» (11.5, 12, 12.5) — приводим к самому частому на странице
    sizes = [round(line["size"] * 2) / 2 for line in lines]
    common = statistics.mode(sizes) if sizes else 0
    for line, size in zip(lines, sizes):
        line["size"] = min(max(common if abs(size - common) <= common * 0.07 else size, 4), 72)
    return {"size": (gray.shape[1], gray.shape[0]), "k": k, "lines": lines, "rules": rules,
            "figures": figures, "chars": sum(len(w["text"]) for line in lines for w in line["words"])}


@functools.lru_cache(maxsize=None)
def ocr_font(bold):
    return pymupdf.Font(fontfile=str(OCR_FONTS[bold]))


def draw_ocr_page(doc, pno, ocr):
    """Вставляет распознанную страницу как обычную PDF-страницу: настоящий текст на тех же местах,
    линии таблиц и картинки. Дальше её, как и остальные страницы, разбирает pdf2docx."""
    k = ocr["k"]
    page = doc.new_page(pno, width=ocr["size"][0] * k, height=ocr["size"][1] * k)
    for box, png in ocr["figures"]:
        page.insert_image(pymupdf.Rect(box) * k, stream=png)
    for x0, y0, x1, y1, thick in ocr["rules"]:
        page.draw_line((x0 * k, y0 * k), (x1 * k, y1 * k), width=max(thick * k, 0.5))
    for bold in (False, True):
        page.insert_font(fontname=f"ocr{int(bold)}", fontfile=str(OCR_FONTS[bold]))
    for line in ocr["lines"]:
        size, font, fontname = line["size"], ocr_font(line["bold"]), f"ocr{int(line['bold'])}"
        base, space = line["base"] * k, font.text_length(" ", line["size"])
        words, x = line["words"], line["words"][0]["box"][0] * k
        for word, after in zip(words, words[1:] + [None]):  # каждое слово — на своё место
            page.insert_text((x, base), word["text"], fontsize=size, fontname=fontname)
            end = x + font.text_length(word["text"], size)
            if not after:
                break
            x = max(after["box"][0] * k, end + space)
            # пробел — посередине промежутка, как при выравнивании по ширине в Word; в широких
            # промежутках (колонки, «должность ____ ФИО») пробела нет: это разные части строки
            if x - end < size * 2:
                page.insert_text(((end + x - space) / 2, base), " ", fontsize=size, fontname=fontname)


# ----------------------------------------------------------------------------- конвертация

def set_text(t, text):
    t.text = text
    t.set(qn("xml:space"), "preserve")


def split_bullets(body):
    """pdf2docx склеивает строки списка в один абзац «• пункт • пункт • пункт» — делим на пункты
    с висячим отступом, как у списков Word."""
    for p in list(body.iter(qn("w:p"))):
        text = Paragraph(p, None).text
        if not text.lstrip().startswith("•") or text.count("•") < 2 or not any(ch.isalnum() for ch in text):
            continue  # не список или колонка одних маркеров (её pdf2docx делает отдельной ячейкой)
        for run in p.findall(qn("w:r")):  # «•» в середине куска текста -> отдельный кусок
            t = run.find(qn("w:t"))
            parts = [part for part in re.split("(?=•)", t.text or "") if part] if t is not None else []
            for part in reversed(parts[1:]):
                piece = copy.deepcopy(run)
                set_text(piece.find(qn("w:t")), part)
                run.addnext(piece)
            if len(parts) > 1:
                set_text(t, parts[0])
        items, started = [p], False
        for run in p.findall(qn("w:r")):
            t = run.find(qn("w:t"))
            if t is not None and (t.text or "").startswith("•"):
                if started:  # новый пункт — новый абзац с тем же оформлением
                    item = copy.deepcopy(p)
                    for child in list(item):
                        if child.tag != qn("w:pPr"):
                            item.remove(child)
                    items[-1].addnext(item)
                    items.append(item)
                started = True
                set_text(t, "•\t" + t.text[1:].lstrip())
            if items[-1] is not p:
                items[-1].append(run)
        for item in items:
            runs = item.findall(qn("w:r"))
            while runs and runs[-1].find(qn("w:br")) is not None and not runs[-1].findall(qn("w:t")):
                item.remove(runs.pop())  # перенос строки в конце пункта больше не нужен
            fmt = Paragraph(item, None).paragraph_format
            fmt.left_indent = Emu((fmt.left_indent or 0) + Pt(18))
            fmt.first_line_indent = Pt(-18)


def polish_docx(doc):
    """Исправляет то, что pdf2docx делает неудобным для Word."""
    body = doc.element.body
    for tbl in body.iter(qn("w:tbl")):  # ширины колонок = ширинам ячеек, иначе колонки «разъезжаются»
        cols = tbl.tblGrid.gridCol_lst
        widths = [None] * len(cols)
        for tr in tbl.tr_lst:
            j = 0
            for tc in tr.tc_lst:
                if tc.grid_span == 1 and tc.width is not None and j < len(widths) and widths[j] is None:
                    widths[j] = tc.width
                j += tc.grid_span
        for col, width in zip(cols, widths):
            if width:
                col.w = width
    for run in body.iter(qn("w:r")):
        fonts = run.find(f"{qn('w:rPr')}/{qn('w:rFonts')}")
        key = "" if fonts is None else (fonts.get(qn("w:ascii")) or "").replace(" ", "").lower()
        texts = [t for t in run.iter(qn("w:t")) if t.text]
        if any(ch in SYMBOL_BULLETS for t in texts for ch in t.text):  # маркер списка «квадратиком»
            for t in texts:
                t.text = "".join(SYMBOL_BULLETS.get(ch, ch) for ch in t.text)
            new_name = "Arial"
        else:
            new_name = FONT_NAMES.get(key.split("-")[0])
        if new_name and fonts is not None:
            for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                if fonts.get(qn(attr)) is not None:
                    fonts.set(qn(attr), new_name)
    split_bullets(body)


def convert(pdf_path, docx_path, lang="rus+eng", dpi=300, force_ocr=False):
    src = pymupdf.open(pdf_path)
    password = None
    if src.needs_pass:
        password = getpass.getpass("PDF защищён паролем. Пароль: ")
        if not src.authenticate(password):
            sys.exit("Неверный пароль.")
    n = len(src)
    scans = [i for i in range(n) if force_ocr or not has_text_layer(src[i])]
    if scans:
        lang = setup_tesseract(lang)
    print(f"Страниц: {n}, с текстом: {n - len(scans)}, сканов (OCR): {len(scans)}")

    bar = tqdm(total=len(scans) * COST_OCR + n * (COST_ANALYZE + COST_PARSE + COST_BUILD),
               desc="Конвертация", bar_format="{desc}: {percentage:3.0f}%|{bar}| {elapsed}<{remaining}{postfix}")

    # 1. Сканы распознаём и заменяем «чистыми» страницами с настоящим текстом
    work = src
    if scans:
        work = pymupdf.open()
        work.insert_pdf(src)
        for i in scans:
            bar.set_postfix_str(f"стр. {i + 1}/{n}: распознавание текста")
            ocr = recognize(src[i], lang, dpi)
            if ocr["chars"] >= MIN_CHARS or force_ocr:  # иначе текста нет (например, фото) — оставляем как есть
                draw_ocr_page(work, i, ocr)
                work.delete_page(i + 1)
            bar.update(COST_OCR)

    # 2. Разбор вёрстки всех страниц через pdf2docx
    cv = Converter(stream=work.tobytes()) if work is not src else Converter(pdf_path, password=password)
    settings = cv.default_settings
    bar.set_postfix_str("анализ структуры документа")
    cv.load_pages().parse_document(**settings)
    bar.update(n * COST_ANALYZE)
    pages, failed = list(cv.pages), []
    for i, page in enumerate(pages):
        bar.set_postfix_str(f"стр. {i + 1}/{n}: разбор вёрстки")
        try:
            page.parse(**settings)
        except Exception as e:
            tqdm.write(f"Стр. {i + 1}: ошибка разбора ({e}), распознаю через OCR")
            failed.append(i)
        bar.update(COST_PARSE)

    if failed:  # страницы, на которых pdf2docx сломался, распознаём как сканы
        lang = setup_tesseract(lang)
        bar.total += len(failed) * COST_OCR
        retry = pymupdf.open()
        for i in failed:
            bar.set_postfix_str(f"стр. {i + 1}/{n}: распознавание текста")
            draw_ocr_page(retry, -1, recognize(src[i], lang, dpi))
            bar.update(COST_OCR)
        cv_retry = Converter(stream=retry.tobytes())
        cv_retry.load_pages().parse_document(**settings)
        for i, page in zip(failed, cv_retry.pages):
            try:
                page.parse(**settings)
                pages[i] = page
            except Exception as e:
                tqdm.write(f"Стр. {i + 1}: не удалось разобрать и после OCR ({e}), страница пропущена")

    # 3. Сборка DOCX
    doc = Document()
    for i, page in enumerate(pages):
        bar.set_postfix_str(f"стр. {i + 1}/{n}: сборка документа")
        if page.finalized:
            try:
                page.make_docx(doc)
            except Exception as e:
                tqdm.write(f"Стр. {i + 1}: не удалось собрать страницу ({e})")
        bar.update(COST_BUILD)
    polish_docx(doc)
    bar.set_postfix_str("сохранение")
    saved = save_docx(doc, Path(docx_path))
    bar.set_postfix_str("готово")
    bar.close()
    cv.close()
    src.close()
    return saved


def save_docx(doc, path):
    """Сохраняет DOCX; если файл открыт в Word — сохраняет под именем «файл (1).docx»."""
    for k in range(100):
        target = path if k == 0 else path.with_name(f"{path.stem} ({k}){path.suffix}")
        try:
            doc.save(target)
            return target
        except PermissionError:
            continue
    raise PermissionError(f"Не удалось сохранить {path}")


def open_file(path):
    """Открывает файл в программе по умолчанию (Word, LibreOffice, ...)."""
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except OSError as e:
        print(f"Не удалось открыть файл автоматически: {e}")


def pick_pdf_windows():
    """Стандартное окно Windows «Открыть файл» (через comdlg32, без tkinter)."""
    import ctypes
    from ctypes import wintypes

    class OPENFILENAMEW(ctypes.Structure):
        _fields_ = [("lStructSize", wintypes.DWORD), ("hwndOwner", wintypes.HWND),
                    ("hInstance", wintypes.HINSTANCE), ("lpstrFilter", ctypes.c_void_p),
                    ("lpstrCustomFilter", ctypes.c_void_p), ("nMaxCustFilter", wintypes.DWORD),
                    ("nFilterIndex", wintypes.DWORD), ("lpstrFile", ctypes.c_void_p),
                    ("nMaxFile", wintypes.DWORD), ("lpstrFileTitle", ctypes.c_void_p),
                    ("nMaxFileTitle", wintypes.DWORD), ("lpstrInitialDir", ctypes.c_void_p),
                    ("lpstrTitle", ctypes.c_void_p), ("Flags", wintypes.DWORD),
                    ("nFileOffset", wintypes.WORD), ("nFileExtension", wintypes.WORD),
                    ("lpstrDefExt", ctypes.c_void_p), ("lCustData", wintypes.LPARAM),
                    ("lpfnHook", ctypes.c_void_p), ("lpTemplateName", ctypes.c_void_p),
                    ("pvReserved", ctypes.c_void_p), ("dwReserved", wintypes.DWORD),
                    ("FlagsEx", wintypes.DWORD)]

    def wide(text):
        return (ctypes.c_wchar * len(text))(*text)

    filters = wide("PDF (*.pdf)\0*.pdf\0Все файлы\0*.*\0\0")
    title = wide("Выберите PDF для конвертации в Word\0")
    path = ctypes.create_unicode_buffer(32768)
    ofn = OPENFILENAMEW(lStructSize=ctypes.sizeof(OPENFILENAMEW),
                        hwndOwner=ctypes.windll.kernel32.GetConsoleWindow(),
                        lpstrFilter=ctypes.addressof(filters), lpstrFile=ctypes.addressof(path),
                        nMaxFile=len(path), lpstrTitle=ctypes.addressof(title),
                        Flags=0x00081808)  # EXPLORER | FILEMUSTEXIST | PATHMUSTEXIST | NOCHANGEDIR
    return path.value if ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(ofn)) else ""


def pick_pdf():
    """Окно выбора PDF (если файл не передан в командной строке)."""
    try:
        if sys.platform == "win32":
            return pick_pdf_windows()
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        path = filedialog.askopenfilename(title="Выберите PDF", filetypes=[("PDF", "*.pdf"), ("Все файлы", "*.*")])
        root.destroy()
        return path
    except Exception:
        return input("Путь к PDF-файлу: ").strip().strip('"')


def main():
    for stream in (sys.stdout, sys.stderr):  # вывод в файл: в Windows иначе cp1252 и кириллица не пишется
        if stream and not stream.isatty():
            stream.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="Конвертер PDF -> DOCX с распознаванием сканов")
    ap.add_argument("pdf", nargs="?", help="PDF-файл (если не указан — откроется окно выбора)")
    ap.add_argument("docx", nargs="?", help="куда сохранить (по умолчанию рядом с PDF)")
    ap.add_argument("--lang", default="rus+eng", help="языки OCR через + (по умолчанию rus+eng)")
    ap.add_argument("--dpi", type=int, default=300, help="разрешение для OCR (по умолчанию 300)")
    ap.add_argument("--ocr", action="store_true", help="распознавать через OCR все страницы")
    ap.add_argument("--no-open", action="store_true", help="не открывать результат после конвертации")
    args = ap.parse_args()

    pdf_path = args.pdf or pick_pdf()
    if not pdf_path:
        sys.exit("Файл не выбран.")
    if not os.path.isfile(pdf_path):
        sys.exit(f"Файл не найден: {pdf_path}")
    docx_path = args.docx or str(Path(pdf_path).with_suffix(".docx"))

    saved = convert(pdf_path, docx_path, args.lang, args.dpi, args.ocr)
    print(f"Готово: {saved}")
    if not args.no_open:
        open_file(str(saved))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nОтменено.")
    except (Exception, SystemExit) as e:
        if len(sys.argv) > 1 or (isinstance(e, SystemExit) and not e.code):
            raise
        # запущен двойным щелчком: не даём окну закрыться, пока ошибку не прочитали
        print(f"Ошибка: {e}")
        input("Нажмите Enter, чтобы закрыть окно...")
        sys.exit(1)
