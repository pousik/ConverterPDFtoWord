#!/usr/bin/env python3
"""
PDF -> Word (DOCX): конвертер с распознаванием сканов (OCR) и прогрессом.

    python pdf2word.py документ.pdf                 -> документ.docx рядом с PDF, сразу открывается
    python pdf2word.py a.pdf b.pdf c.pdf            -> каждый в свой .docx (можно перетащить файлы на exe)
    python pdf2word.py документ.pdf -o итог.docx    -> свой путь (или -o папка)
    python pdf2word.py                              -> откроется окно выбора файлов

Страницы с текстовым слоем конвертирует pdf2docx (сохраняются шрифты, таблицы,
картинки). Страницы-сканы распознаёт Tesseract OCR: текст, линии таблиц и картинки
переносятся на «чистую» PDF-страницу, которую затем так же разбирает pdf2docx —
получаются обычные редактируемые абзацы, таблицы и рисунки. Существующие файлы
никогда не перезаписываются: если имя занято, результат получит имя «файл (1).docx».
"""
import argparse
import copy
import functools
import getpass
import glob
import itertools
import logging
import os
import re
import shutil
import statistics
import subprocess
import sys
import tempfile
import threading
import traceback
import xml.etree.ElementTree as ET
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from pathlib import Path

import cv2
import numpy as np
import pymupdf
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Emu, Pt
from docx.text.paragraph import Paragraph
from tqdm import tqdm

# служебные сообщения PyMuPDF и INFO-логи pdf2docx ломают прогресс-бар — прячем их
pymupdf.set_messages(pylogging=True, pylogging_level=logging.DEBUG)
from pdf2docx import Converter  # noqa: E402
from pdf2docx.text.Line import Line  # noqa: E402
from pdf2docx.text.TextSpan import TextSpan  # noqa: E402
logging.getLogger().setLevel(logging.ERROR)

APP_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))  # папка программы / распакованного exe
BUNDLED_TESSERACT = APP_DIR / "tesseract"  # Tesseract, встроенный в PDF2Word.exe (см. PDF2Word.spec)
FONT_FILES = {  # шрифт распознанного текста: свой (метрики как у Times New Roman), иначе системный
    False: [APP_DIR / "assets" / "fonts" / "LiberationSerif-Regular.ttf",
            Path(sys.executable).parent / "assets" / "fonts" / "LiberationSerif-Regular.ttf",
            Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / "times.ttf",
            Path("/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf"),
            Path("/usr/share/fonts/liberation-serif/LiberationSerif-Regular.ttf"),
            Path("/System/Library/Fonts/Supplemental/Times New Roman.ttf")],
    True: [APP_DIR / "assets" / "fonts" / "LiberationSerif-Bold.ttf",
           Path(sys.executable).parent / "assets" / "fonts" / "LiberationSerif-Bold.ttf",
           Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / "timesbd.ttf",
           Path("/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf"),
           Path("/usr/share/fonts/liberation-serif/LiberationSerif-Bold.ttf"),
           Path("/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf")],
}
MIN_CHARS = 20          # меньше букв/цифр на странице -> считаем её сканом
MAX_PIXELS = 40_000_000  # предел картинки страницы для OCR (~A2 при 300 dpi): плакаты рендерятся с меньшим dpi
WORKERS = min(os.cpu_count() or 1, 8)  # сканы распознаются параллельно, по странице на ядро
TESSERACT_ARGS = ["--oem", "1", "-c", "tessedit_do_invert=0"]  # LSTM; белый текст на чёрном не ищем
SIZE_FACTOR = 1.08      # x_size строки от Tesseract -> кегль шрифта (подобрано на тестах)
XHTML = "{http://www.w3.org/1999/xhtml}"
IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".gif", ".webp", ".jp2", ".jpx"}
# «Вес» шагов в прогресс-баре, чтобы проценты шли примерно равномерно по времени
COST_OCR, COST_ANALYZE, COST_PARSE, COST_BUILD = 25, 1, 2, 1

# маркеры списков из шрифтов Symbol/Wingdings (частная область Unicode) -> обычные символы
SYMBOL_BULLETS = {"symbol": {"\uf0b7": "•"},
                  "wingdings": {"\uf0a7": "▪", "\uf0d8": "➢", "\uf076": "❖", "\uf0fc": "✓",
                                "\uf0e8": "➔", "\uf06e": "■", "\uf071": "❑", "\uf09f": "•"}}
# свободные шрифты, которыми PDF часто заменяет шрифты Windows -> их оригиналы
FONT_NAMES = {"liberationserif": "Times New Roman", "tinos": "Times New Roman", "nimbusroman": "Times New Roman",
              "liberationsans": "Arial", "arimo": "Arial", "nimbussans": "Arial",
              "liberationmono": "Courier New", "cousine": "Courier New"}
# так Tesseract нередко читает маркер списка «•»
BULLET_LIKE = {"•", "·", "●", "®", "©", "*", "»", "«", "°", "¢", "e", "o", "е", "о"}
# дефис в конце строки, который нельзя убирать при склейке переносов: «кто-|нибудь», «северо-|западный»
HYPHEN_KEEP_AFTER = {"то", "либо", "нибудь", "ка", "таки", "де"}
HYPHEN_KEEP_BEFORE = {"кое", "кой", "северо", "юго", "вице", "экс", "обер", "унтер", "штаб", "контр", "лейб"}
HYPHEN_KEEP_PAIRS = {("из", "за"), ("из", "под"), ("по", "за")}
INVALID_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff\ud800-\udfff]")  # ломают сохранение DOCX

WINDOWS_TESSERACT = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
    os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe"),
]


class UserError(Exception):
    """Понятная пользователю ошибка: печатается без трассировки."""


class OcrUnavailable(UserError):
    """Tesseract не найден или без нужных языков: сканы остаются картинками."""


WARNINGS = []  # предупреждения за запуск (например, «страница вставлена картинкой») — их надо успеть прочитать


def warn(message):
    WARNINGS.append(message)
    tqdm.write(message)


def can_join(left, right):
    """Можно ли склеить перенос «left-» + «right» в одно слово (а не «кто-нибудь» -> «ктонибудь»)."""
    left_word = re.split(r"[^\w]", left)[-1].lower()
    right_word = re.split(r"[^\w-]", right)[0].lower()
    return (len(left_word) > 1 and left_word[-1:].isalpha() and right[:1].islower()
            and right_word.split("-")[0] not in HYPHEN_KEEP_AFTER and left_word not in HYPHEN_KEEP_BEFORE
            and (left_word, right_word) not in HYPHEN_KEEP_PAIRS)


# ----------------------------------------------------------------------------- поправки pdf2docx

def _make_line_docx(self, p, _make_line=Line.make_docx):
    """pdf2docx склеивает строки абзаца без пробела («Покупатель» + «обязуется» -> «Покупательобязуется»),
    если в PDF строка не заканчивается пробелом, и оставляет переносы «уполномочен-ными». Чиним оба."""
    _make_line(self, p)
    if self.line_break:
        return
    texts = [t for t in p._p.iter(qn("w:t")) if t.text]  # последний текст, в том числе внутри гиперссылки
    if not texts or texts[-1].text.endswith((" ", "\u00ad", "\t")):
        return
    last = texts[-1]
    if last.text.endswith("-"):
        lines = list(getattr(self.parent, "lines", []))
        index = next((i for i, line in enumerate(lines) if line is self), None)
        after = lines[index + 1].text.lstrip() if index is not None and index + 1 < len(lines) else ""
        if after and can_join(last.text[:-1], after):
            last.text = last.text[:-1]  # перенос слова: дефис убираем, следующая строка продолжит слово
        return
    last.text += " "
    last.set(qn("xml:space"), "preserve")


def _make_span_docx(self, paragraph, _make_span=TextSpan.make_docx):
    """Символы, недопустимые в XML (U+FFFE, управляющие — из «битых» шрифтов), ломают сборку страницы."""
    if self.chars and any(INVALID_XML.search(c.c) for c in self.chars):
        self.chars = [c for c in self.chars if not INVALID_XML.search(c.c)]
    elif not self.chars and self._text and INVALID_XML.search(self._text):
        self._text = INVALID_XML.sub("", self._text)
    _make_span(self, paragraph)


Line.make_docx = _make_line_docx
TextSpan.make_docx = _make_span_docx


# ----------------------------------------------------------------------------- Tesseract

class TesseractError(RuntimeError):
    pass


TESSERACT = None  # путь к tesseract(.exe); задаётся в setup_tesseract
TESS_CWD = None   # папка запуска Tesseract: языки ищутся относительно неё (так не мешает кириллица в пути)
TESS_ENV = None
TESS_PROCS, TESS_LOCK = set(), threading.Lock()  # запущенные Tesseract — чтобы прервать их по Ctrl+C


def setup_tesseract(lang):
    """Находит Tesseract и оставляет только установленные языки. Без него — OcrUnavailable."""
    global TESSERACT, TESS_CWD, TESS_ENV
    env = dict(os.environ, OMP_THREAD_LIMIT="1")  # страницы распознаются параллельно, сам Tesseract — в 1 поток
    if (BUNDLED_TESSERACT / "tesseract.exe").is_file():
        exe = BUNDLED_TESSERACT / "tesseract.exe"
    else:
        found = shutil.which("tesseract") or next((p for p in WINDOWS_TESSERACT if os.path.isfile(p)), None)
        exe = Path(found) if found else None
    if exe is None:
        raise OcrUnavailable("Tesseract OCR не найден — сканы будут вставлены в документ картинками.\n"
                             "Установите его: https://github.com/UB-Mannheim/tesseract/wiki (отметьте язык Russian).")
    cwd = None
    if (exe.parent / "tessdata").is_dir():
        # Tesseract читает путь к языкам в кодировке ANSI: абсолютный путь с кириллицей ломается,
        # поэтому запускаем его из его папки и передаём относительный путь
        cwd, env["TESSDATA_PREFIX"] = str(exe.parent), "tessdata"
    try:
        listing = subprocess.run([str(exe), "--list-langs"], capture_output=True, cwd=cwd, env=env,
                                 creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except OSError as e:
        raise OcrUnavailable(f"Не удаётся запустить Tesseract ({exe}): {e}") from e
    output = listing.stdout.decode("utf-8", "replace")
    if listing.returncode:
        raise OcrUnavailable(f"Tesseract ({exe}) не работает: {listing.stderr.decode('utf-8', 'replace').strip()}")
    installed = {code.strip() for code in output.splitlines()[1:]} - {""}
    langs = [code for code in lang.split("+") if code in installed]
    missing = [code for code in lang.split("+") if code and code not in installed]
    if not langs:
        raise OcrUnavailable(f"В Tesseract нет языков «{lang}» (есть: {', '.join(sorted(installed)) or 'ни одного'}). "
                             "Переустановите его, отметив язык Russian.")
    if missing:
        warn(f"Внимание: языки OCR не установлены и будут пропущены: {', '.join(missing)}")
    TESSERACT, TESS_CWD, TESS_ENV = str(exe), cwd, env
    return "+".join(langs)


def tesseract(img, lang, *args):
    """Запускает Tesseract на картинке из памяти: через stdin, без временных файлов и сжатия в PNG."""
    proc = subprocess.Popen([TESSERACT, "stdin", "stdout", "-l", lang, *args], cwd=TESS_CWD, env=TESS_ENV,
                            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    with TESS_LOCK:
        TESS_PROCS.add(proc)
    try:
        out, err = proc.communicate(cv2.imencode(".bmp", img)[1].tobytes())
    finally:
        with TESS_LOCK:
            TESS_PROCS.discard(proc)
    if proc.returncode:
        raise TesseractError(err.decode("utf-8", "replace").strip() or f"код выхода {proc.returncode}")
    return out.decode("utf-8", "replace")


def stop_tesseract():
    """Прерывает все запущенные Tesseract (Ctrl+C, ошибка)."""
    with TESS_LOCK:
        for proc in TESS_PROCS:
            try:
                proc.kill()
            except OSError:
                pass


# ----------------------------------------------------------------------------- анализ страниц

def visible_text(page):
    """Видимые символы страницы (без невидимого OCR-слоя и прозрачного текста)."""
    try:
        spans = [s for s in page.get_texttrace() if s["type"] != 3 and s.get("opacity", 1) > 0]
        return "".join(chr(c[0]) for s in spans for c in s["chars"] if c[0] > 0)
    except Exception:  # редкие «битые» шрифты
        return page.get_text()


def has_text_layer(page):
    """True, если на странице есть видимый и читаемый текст, а не только картинка-скан. Не считаются:
    невидимый OCR-слой, текст, спрятанный под сканом, и мелкие надписи поверх скана (штамп ЭЦП и т.п.)."""
    chars = visible_text(page)
    good = sum(ch.isalnum() for ch in chars)
    if good < MIN_CHARS or chars.count("\ufffd") >= good:
        return False
    area, rect = abs(page.rect) or 1, page.rect
    try:
        log = [(kind, pymupdf.Rect(box)) for kind, box in page.get_bboxlog()]
    except Exception:
        return True
    scans = [i for i, (kind, box) in enumerate(log) if kind == "fill-image" and abs(box & rect) > 0.5 * area]
    if not scans:
        return True
    last, scan = scans[-1], log[scans[-1]][1]
    shown = sum(abs(box & rect) for i, (kind, box) in enumerate(log) if kind in ("fill-text", "stroke-text")
                and (i > last or abs(box & scan) < 0.5 * abs(box)))  # поверх скана или вне его
    return shown > 0.05 * area


def render_dpi(page, dpi):
    """(dpi рендера, dpi для OCR и размеров). Огромные страницы рендерятся с меньшим dpi (память), а фото,
    у которых размер страницы взят из пикселей «по 72 dpi» (сканеры в телефоне), — в их родном разрешении."""
    w, h = max(page.rect.width / 72, 0.1), max(page.rect.height / 72, 0.1)
    limit = min((MAX_PIXELS / (w * h)) ** 0.5, 30000 / max(w, h))
    if max(w, h) > 17:  # больше A3 — проверим, не фото ли это на всю страницу
        try:
            native = [info["width"] * 72 / max(info["bbox"][2] - info["bbox"][0], 1) for info in page.get_image_info()
                      if abs(pymupdf.Rect(info["bbox"]) & page.rect) > 0.5 * abs(page.rect)]
        except Exception:
            native = []
        if native:  # пиксели фото как у обычного скана: размеры и OCR считаем «по 300 dpi»
            return max(min(dpi, limit, max(native)), 10), dpi
    r = max(min(dpi, limit), 10)
    return r, r


def render(page, dpi):
    """Картинка страницы (PyMuPDF не потокобезопасен, поэтому рендер — только в основном потоке)."""
    pix = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csRGB, alpha=False)
    return np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, 3)


def hocr_value(el, key):
    """Числа из title-атрибута hOCR: 'bbox 10 20 30 40; x_size 47' -> [10, 20, 30, 40]."""
    m = re.search(rf"{key} ([^;]+)", el.get("title", ""))
    try:
        return [float(v) for v in m.group(1).split()] if m else None
    except ValueError:
        return None


def rotate(img, angle, border=None):
    """Поворачивает изображение на angle градусов против часовой; края — продолжением фона (или border)."""
    h, w = img.shape[:2]
    matrix = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1)
    if border is None:
        return cv2.warpAffine(img, matrix, (w, h), flags=cv2.INTER_LINEAR, borderMode=cv2.BORDER_REPLICATE)
    return cv2.warpAffine(img, matrix, (w, h), flags=cv2.INTER_LINEAR, borderValue=(border,) * 3)


def text_ink(small):
    """«Чернила» на уменьшенной картинке: темнее местного фона (работает и на серой бумаге)."""
    return cv2.adaptiveThreshold(small, 1, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 15, 12)


def osd_rotation(gray):
    """На сколько градусов по часовой повернуть страницу (OSD Tesseract, по уменьшенной копии)."""
    half = cv2.resize(gray, None, fx=0.5, fy=0.5, interpolation=cv2.INTER_AREA)
    try:
        rotate_by = int(re.search(r"Rotate: (\d+)", tesseract(half, "osd", "--psm", "0", "--dpi", "150")).group(1))
        return rotate_by if rotate_by in (90, 180, 270) else 0
    except (TesseractError, AttributeError, ValueError):
        return 0  # слишком мало текста для определения ориентации


def ink_profiles(small):
    """Неравномерность «чернил» по строкам и по столбцам: у строк текста профиль «полосатый»."""
    ink = text_ink(small).astype(np.float32)
    rows, cols = ink.sum(axis=1), ink.sum(axis=0)
    return float(rows.std() / (rows.mean() + 1e-6)), float(cols.std() / (cols.mean() + 1e-6))


def skew_angle(gray):
    """Наклон скана: при верном угле строки текста дают самые резкие «пики» в сумме по строкам."""
    def best(img, angles):
        ink = text_ink(img).astype(np.float32)
        return float(max(angles, key=lambda a: (float(np.var(rotate(ink, a, border=0).sum(axis=1))), -abs(a))))

    tiny = cv2.resize(gray, None, fx=0.125, fy=0.125, interpolation=cv2.INTER_AREA)
    small = cv2.resize(gray, None, fx=0.25, fy=0.25, interpolation=cv2.INTER_AREA)
    if min(tiny.shape) < 16:  # слишком маленькая картинка, чтобы судить о наклоне
        return 0.0
    coarse = best(tiny, np.arange(-5, 5.01, 0.5))
    return best(small, np.arange(coarse - 0.5, coarse + 0.51, 0.1))


def flatten_background(gray):
    """Серая, жёлтая или неравномерно освещённая бумага -> белая: делим на оценку фона. Для белой бумаги
    картинка не меняется (так распознавание обычных сканов остаётся прежним)."""
    small = cv2.resize(gray, None, fx=0.125, fy=0.125, interpolation=cv2.INTER_AREA)
    background = cv2.GaussianBlur(cv2.dilate(small, np.ones((9, 9), np.uint8)), (0, 0), 3)
    if np.percentile(background, 5) >= 215:
        return gray
    background = cv2.resize(background, (gray.shape[1], gray.shape[0]), interpolation=cv2.INTER_LINEAR)
    return cv2.divide(gray, np.maximum(background, 1), scale=255)


def stroke_width(ink, box):
    """Толщина штриха букв: медиана длин горизонтальных отрезков «чернил» в рамке строки."""
    x0, y0, x1, y1 = (max(int(v), 0) for v in box)
    region = ink[y0:y1, x0:x1]
    if region.size == 0:
        return 0.0
    edges = np.diff(np.pad(region, ((0, 0), (1, 1))).astype(np.int8), axis=1)
    runs = np.nonzero(edges == -1)[1] - np.nonzero(edges == 1)[1]
    return float(np.median(runs)) if runs.size else 0.0


def find_rules(ink, dpi):
    """Линии таблиц и подчёркивания: длинные тонкие горизонтальные и вертикальные отрезки (в пикселях)."""
    rules, mask = [], np.zeros_like(ink)
    for kernel in ((max(dpi // 2, 2), 1), (1, max(dpi * 3 // 10, 2))):
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
        if len(last["text"]) > 2 and last["text"].endswith("-") and can_join(last["text"][:-1], first["text"]):
            last["text"] = last["text"][:-1] + first["text"]
            b["words"].pop(0)
    return [line for line in lines if line["words"]]


def words_box(words):
    """Общая рамка нескольких слов."""
    return [min(w["box"][0] for w in words), min(w["box"][1] for w in words),
            max(w["box"][2] for w in words), max(w["box"][3] for w in words)]


def mask_words(mask, lines, pad_x=4, pad_y=4):
    """Закрашивает на маске рамки всех слов (с запасом)."""
    for line in lines:
        for word in line["words"]:
            x0, y0, x1, y1 = (int(v) for v in word["box"])
            mask[max(y0 - pad_y, 0):max(y1 + pad_y, 0), max(x0 - pad_x, 0):max(x1 + pad_x, 0)] = 0


def recover_missed(clean, ink, rules_mask, lines, lang, dpi):
    """Второй проход: дочитывает то, что Tesseract пропустил (одиночные цифры в ячейках таблиц, номера
    страниц), и перечитывает короткие слова, в которых он не уверен. Все кусочки собираются полосками
    на одну картинку и распознаются за один запуск."""
    if not lines:
        return []
    x_size = max(int(statistics.median(line["x_size"] for line in lines)), 4)
    leftover = ink.copy()
    leftover[rules_mask > 0] = 0
    mask_words(leftover, lines, pad_x=x_size // 2 + 4)  # вплотную к словам — «хвосты» распознанных букв (№, Р)
    merged = cv2.dilate(leftover, np.ones((3, max(x_size * 3 // 5, 1)), np.uint8))  # буквы слова — в кусок
    crops, pad = [], x_size // 3  # (x, y, картинка, слово для перечитывания или None)
    for x, y, w, h, _ in cv2.connectedComponentsWithStats(merged)[2][1:]:
        if 0.4 * x_size <= h <= 1.8 * x_size and w <= dpi * 2 and len(crops) < 50:
            x0, y0 = max(x - pad, 0), max(y - pad, 0)
            crops.append((x0, y0, clean[y0:y + h + pad, x0:x + w + pad], None))
    for line in lines:
        words = line["words"]
        for j, word in enumerate(words):  # короткое слово «на отшибе» (ячейка таблицы, номер), Tesseract не уверен
            alone = (j == 0 or word["box"][0] - words[j - 1]["box"][2] > x_size) and \
                    (j == len(words) - 1 or words[j + 1]["box"][0] - word["box"][2] > x_size)
            if alone and len(word["text"]) <= 3 and word["text"] != "•" and word["conf"] < 80 and len(crops) < 100:
                x0, y0, x1, y1 = (int(v) for v in word["box"])
                x0, y0 = max(x0 - 6, 0), max(y0 - pad, 0)
                crops.append((x0, y0, clean[y0:y1 + pad, x0:x1 + 6], word))
    crops = [c for c in crops if c[2].size]
    if not crops:
        return []

    canvas = np.full((sum(c[2].shape[0] + x_size for c in crops) + x_size,
                      max(c[2].shape[1] for c in crops) + 2 * x_size), 255, np.uint8)
    bands, top = [], x_size
    for x0, y0, crop, _ in crops:
        canvas[top:top + crop.shape[0], x_size:x_size + crop.shape[1]] = crop
        bands.append((top, top + crop.shape[0], x0 - x_size, y0 - top))
        top += crop.shape[0] + x_size
    tsv = tesseract(canvas, lang, *TESSERACT_ARGS, "--psm", "6", "--dpi", str(int(dpi)), "tsv")
    found = {}
    for row in tsv.splitlines()[1:]:
        cells = row.split("\t")
        if len(cells) < 12:
            continue
        try:
            x, y, width, height = (int(v) for v in cells[6:10])
            text, conf = cells[11].strip(), float(cells[10])
        except ValueError:
            continue
        if not text or conf < 70 or not any(ch.isalnum() for ch in text):
            continue
        middle = y + height / 2
        for band, (y_top, y_bottom, dx, dy) in enumerate(bands):
            if y_top <= middle <= y_bottom:
                box = [x + dx, y + dy, x + width + dx, y + height + dy]
                found.setdefault(band, []).append({"box": box, "text": text, "conf": conf})

    dark, new_lines = clean < 128, []
    for band, words in found.items():
        target = crops[band][3]
        if target is None:  # пропущенный кусок текста — новая строка
            new_lines.append({"words": sorted(words, key=lambda w: w["box"][0]), "par": ("missed", band),
                              "base": max(w["box"][3] for w in words), "x_size": x_size,
                              "stroke": stroke_width(dark, words_box(words)) / x_size})
        elif len(words) == 1 and words[0]["conf"] > target["conf"] + 10:  # перечитанное слово увереннее
            target["text"], target["conf"] = words[0]["text"], words[0]["conf"]
    return new_lines


def parse_hocr(hocr, dark):
    """Строки из hOCR: слова (рамка, текст, уверенность), базовая линия, высота строки, толщина штриха."""
    lines = []
    try:
        root = ET.fromstring(hocr)
    except ET.ParseError:
        return lines
    for n, par in enumerate(root.iter(f"{XHTML}p")):
        for line in par:
            words = []
            for w in line:
                text = INVALID_XML.sub("", "".join(w.itertext())).strip()
                conf = (hocr_value(w, "x_wconf") or [100])[0]
                box = hocr_value(w, "bbox")
                if text and box and len(box) == 4 and box[2] > box[0] and box[3] > box[1] \
                        and (conf >= 30 or any(ch.isalnum() for ch in text)):  # иначе это грязь скана
                    words.append({"box": box, "text": text, "conf": conf})
            box = hocr_value(line, "bbox")
            if not words or not box or len(box) != 4:
                continue
            x_size = max((hocr_value(line, "x_size") or [box[3] - box[1]])[0], 4)
            if len(words) > 2 and words[0]["text"] in BULLET_LIKE:  # маркер: отделён заметно больше, чем слова
                gaps = [b["box"][0] - a["box"][2] for a, b in zip(words, words[1:])]
                if gaps[0] > 0.8 * x_size and gaps[0] > 2 * statistics.median(gaps[1:]):
                    words[0]["text"] = "•"
            elif len(words) == 2 and words[0]["text"] in BULLET_LIKE \
                    and words[1]["box"][0] - words[0]["box"][2] > 0.8 * x_size:
                words[0]["text"] = "•"
            baseline = hocr_value(line, "baseline") or [0, 0]
            lines.append({"words": words, "par": n, "base": box[3] + (baseline[1] if len(baseline) > 1 else 0),
                          "x_size": x_size, "stroke": stroke_width(dark, box) / x_size})
    return lines


def find_figures(gray, color, rest, rules_mask, edge, dpi):
    """Картинки (фото, графики, печати, подписи): связные области «не-бумаги» вне слов и линий.
    Отбрасываются равномерные пятна (тени, виньетирование, заливка ячеек таблицы)."""
    h, w = gray.shape
    blobs = cv2.dilate(rest, np.ones((25, 25), np.uint8))
    boxes = [[x, y, x + bw, y + bh] for x, y, bw, bh, _ in cv2.connectedComponentsWithStats(blobs)[2][1:]
             if min(bw, bh) >= dpi // 10]
    gap, merged = int(dpi // 2), True
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
    edges = cv2.Canny(gray, 40, 120)
    figures, band = [], edge + 20
    for x0, y0, x1, y1 in boxes:
        x0, y0, x1, y1 = max(x0, 0), max(y0, 0), min(x1, w), min(y1, h)
        if x1 - x0 < dpi * 0.4 or y1 - y0 < dpi * 0.3 or rest[y0:y1, x0:x1].mean() <= 255 * 0.03:
            continue
        detail = edges[y0:y1, x0:x1].mean() / 255
        at_border = x0 <= band or y0 <= band or x1 >= w - band or y1 >= h - band
        if detail < 0.002 or (at_border and detail < 0.01):
            continue  # ровное пятно: тень сканера, виньетирование, серый фон
        sides = [rules_mask[max(y0 + 12 - 15, 0):y0 + 12 + 15, x0:x1], rules_mask[max(y1 - 12 - 15, 0):y1 - 12 + 15, x0:x1],
                 rules_mask[y0:y1, max(x0 + 12 - 15, 0):x0 + 12 + 15], rules_mask[y0:y1, max(x1 - 12 - 15, 0):x1 - 12 + 15]]
        lined = sum(side.size and (side.max(axis=0 if i < 2 else 1) > 0).mean() > 0.6 for i, side in enumerate(sides))
        if lined >= 3:
            continue  # закрашенная ячейка или строка таблицы
        crop = cv2.cvtColor(color[y0:y1, x0:x1], cv2.COLOR_RGB2BGR)
        figures.append(((x0, y0, x1, y1), cv2.imencode(".png", crop)[1].tobytes()))
    return figures


def recognize(color, lang, dpi, check_turn=True):
    """Распознаёт страницу-скан: строки слов с кеглем и жирностью, линии таблиц и картинки."""
    dpi = int(dpi)
    gray = cv2.medianBlur(cv2.cvtColor(color, cv2.COLOR_RGB2GRAY), 3)  # убирает «пыль» скана
    if min(gray.shape) < 32:  # крошечная страница — распознавать нечего
        return {"size": (gray.shape[1], gray.shape[0]), "lines": [], "rules": [], "figures": [], "chars": 0}
    rows, cols = ink_profiles(cv2.resize(gray, None, fx=0.25, fy=0.25, interpolation=cv2.INTER_AREA))
    if check_turn and cols > rows * 1.2:  # строки идут вертикально — похоже, страница отсканирована боком
        turn = osd_rotation(gray)
        if turn:
            return recognize(turn_image(color, turn), lang, dpi, check_turn=False)
    angle = skew_angle(gray)
    if abs(angle) > 0.05:  # выравниваем наклон: строки и линии таблиц становятся строго горизонтальными
        color, gray = rotate(color, angle), rotate(gray, angle)
    gray = flatten_background(gray)

    ink = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 51, 20)
    edge = max(int(dpi * 0.15), 1)  # тени и края листа у границы скана — не содержимое
    ink[:edge], ink[-edge:], ink[:, :edge], ink[:, -edge:] = 0, 0, 0, 0
    rules, rules_mask = find_rules(ink, dpi)
    clean = gray.copy()
    clean[rules_mask > 0] = 255  # без линий Tesseract намного лучше читает текст в ячейках таблиц
    dark = clean < 128

    lines = parse_hocr(tesseract(clean, lang, *TESSERACT_ARGS, "--psm", "3", "--dpi", str(dpi), "hocr"), dark)
    if sum(len(w["text"]) for line in lines for w in line["words"]) < MIN_CHARS \
            and (ink[rules_mask == 0] > 0).mean() > 0.002:  # мало текста, но «чернила» есть: таблица цифр и т.п.
        sparse = parse_hocr(tesseract(clean, lang, *TESSERACT_ARGS, "--psm", "6", "--dpi", str(dpi), "hocr"), dark)
        if sum(len(w["text"]) for line in sparse for w in line["words"]) > \
                sum(len(w["text"]) for line in lines for w in line["words"]):
            lines = sparse

    confs = [w["conf"] for line in lines for w in line["words"] if any(ch.isalnum() for ch in w["text"])]
    if check_turn and len(confs) >= 10 and statistics.mean(confs) < 55:  # «каша» — может, вверх ногами?
        turn = osd_rotation(gray)
        if turn:
            return recognize(turn_image(color, turn), lang, dpi, check_turn=False)
    lines += recover_missed(clean, ink, rules_mask, lines, lang, dpi)

    # жирный шрифт: штрихи заметно толще, чем у основного текста страницы
    normal = statistics.median(line["stroke"] for line in lines) if lines else 0
    for line in lines:
        line["bold"] = normal > 0 and line["stroke"] > normal * 1.3

    paper = float(np.percentile(gray, 95))
    rest = np.where(gray < paper - 45, 255, 0).astype(np.uint8)  # заметно темнее бумаги
    rest[:edge], rest[-edge:], rest[:, :edge], rest[:, -edge:] = 0, 0, 0, 0
    rest[rules_mask > 0] = 0
    mask_words(rest, lines)
    figures = find_figures(gray, color, rest, rules_mask, edge, dpi)

    def in_figure(box):  # «слова», которые Tesseract нашёл внутри картинки, обычно мусор
        cx, cy = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
        return any(f[0] <= cx <= f[2] and f[1] <= cy <= f[3] for f, _ in figures)

    for line in lines:
        line["words"] = [w for w in line["words"] if w["conf"] >= 70 or not in_figure(w["box"])]
    rules = [r for r in rules if not in_figure(r[:4])]  # края столбиков графика — не таблица
    # переносы склеиваем в самом конце, когда все кусочки уже дочитаны (иначе «хвост» прочтётся дважды)
    lines = [line for _, group in itertools.groupby(lines, key=lambda line: line["par"])
             for line in join_hyphenated([line for line in group if line["words"]])]
    return {"size": (gray.shape[1], gray.shape[0]), "lines": lines, "rules": rules, "figures": figures,
            "chars": sum(len(w["text"]) for line in lines for w in line["words"])}


def turn_image(img, turn):
    """Поворот на turn градусов по часовой (кратно 90)."""
    return np.ascontiguousarray(np.rot90(img, -turn // 90)) if turn else img


def recognize_safe(color, lang, dpi):
    """recognize, но ошибка одной страницы не обрывает весь документ: возвращается как результат."""
    try:
        return recognize(color, lang, dpi)
    except Exception as e:  # noqa: BLE001 — страница останется картинкой, остальные распознаются
        return e


def recognize_pages(src, pages, lang, dpi, on_done):
    """Распознаёт страницы параллельно: рендер — в основном потоке, остальное — в рабочих.
    Одновременно в памяти не больше картинок, чем рабочих потоков. on_done(номер, результат или
    исключение, dpi размеров) вызывается в основном потоке сразу по готовности страницы."""
    queue, running = iter(pages), {}
    pool = ThreadPoolExecutor(WORKERS)
    try:
        def submit():
            for i in queue:
                try:
                    render_at, size_dpi = render_dpi(src[i], dpi)
                    running[pool.submit(recognize_safe, render(src[i], render_at), lang, size_dpi)] = (i, size_dpi)
                    return
                except Exception as e:  # noqa: BLE001 — не удалось даже отрисовать страницу
                    on_done(i, e, dpi)

        for _ in range(WORKERS):
            submit()
        while running:
            for future in wait(running, timeout=0.5, return_when=FIRST_COMPLETED)[0]:
                i, size_dpi = running.pop(future)
                submit()  # следующую страницу — в работу сразу, до обработки готовой
                on_done(i, future.result(), size_dpi)
    except BaseException:
        stop_tesseract()
        pool.shutdown(wait=False, cancel_futures=True)
        raise
    pool.shutdown()


# ----------------------------------------------------------------------------- «чистая» страница

@functools.lru_cache(maxsize=None)
def ocr_font(bold):
    """Шрифт распознанного текста (с кириллицей и метриками Times New Roman)."""
    for path in FONT_FILES[bold]:
        if path.is_file():
            return pymupdf.Font(fontfile=str(path))
    raise UserError("Не найден шрифт для распознанного текста (assets/fonts/LiberationSerif-*.ttf или Times New Roman). "
                    "Соберите программу через PDF2Word.spec — он кладёт шрифты в сборку.")


def assign_sizes(ocr, k):
    """Кегль строк: по ширине слов (шрифт метрически как Times New Roman), с проверкой по высоте строки."""
    lines = ocr["lines"]
    for line in lines:
        by_height = line["x_size"] * k * SIZE_FACTOR
        natural = sum(ocr_font(line["bold"]).text_length(w["text"], 1) for w in line["words"])
        by_width = sum(w["box"][2] - w["box"][0] for w in line["words"]) * k / natural if natural else 0
        line["size"] = by_width if by_height and 0.75 < by_width / by_height < 1.35 else by_height
    # кегль основного текста немного «гуляет» (11.5, 12, 12.5) — приводим к самому частому на странице
    sizes = [round(line["size"] * 2) / 2 for line in lines]
    common = statistics.mode(sizes) if sizes else 0
    for line, size in zip(lines, sizes):
        line["size"] = min(max(common if abs(size - common) <= common * 0.07 else size, 4), 72)


def draw_ocr_page(doc, pno, ocr, size_dpi):
    """Вставляет распознанную страницу как обычную PDF-страницу: настоящий текст на тех же местах,
    линии таблиц и картинки. Дальше её, как и остальные страницы, разбирает pdf2docx."""
    k = 72 / size_dpi  # пиксели -> пункты
    assign_sizes(ocr, k)
    page = doc.new_page(pno, width=max(ocr["size"][0] * k, 1), height=max(ocr["size"][1] * k, 1))
    for box, png in ocr["figures"]:
        page.insert_image(pymupdf.Rect(box) * k, stream=png)
    for x0, y0, x1, y1, thick in ocr["rules"]:
        page.draw_line((x0 * k, y0 * k), (x1 * k, y1 * k), width=max(thick * k, 0.5))
    writer = pymupdf.TextWriter(page.rect)  # весь текст страницы — одной операцией, это быстро
    for line in ocr["lines"]:
        size, font = line["size"], ocr_font(line["bold"])
        base, space = line["base"] * k, font.text_length(" ", line["size"])
        words, x = line["words"], line["words"][0]["box"][0] * k
        for word, after in zip(words, words[1:] + [None]):  # каждое слово — на своё место
            writer.append((x, base), word["text"], font=font, fontsize=size)
            end = x + font.text_length(word["text"], size)
            if not after:
                break
            x = max(after["box"][0] * k, end + space)
            # пробел — посередине промежутка, как при выравнивании по ширине в Word; в широких
            # промежутках (колонки, «должность ____ ФИО») пробела нет: это разные части строки
            if x - end < size * 2:
                writer.append(((end + x - space) / 2, base), " ", font=font, fontsize=size)
    writer.write_text(page)


def image_page(doc, page, dpi=150):
    """Последний рубеж: страница, которую не удалось ни разобрать, ни распознать, — картинкой."""
    pix = page.get_pixmap(dpi=dpi, alpha=False)
    out = doc.new_page(-1, width=page.rect.width, height=page.rect.height)
    out.insert_image(out.rect, stream=pix.tobytes("png"))


# ----------------------------------------------------------------------------- DOCX

def set_text(t, text):
    t.text = text
    t.set(qn("xml:space"), "preserve")


def split_bullets(body):
    """pdf2docx склеивает строки списка в один абзац «• пункт • пункт • пункт» — делим на пункты
    с висячим отступом, как у списков Word. Делим только там, где кусок текста начинается с «•»."""
    for p in list(body.iter(qn("w:p"))):
        text = Paragraph(p, None).text
        if not text.lstrip().startswith("•") or not any(ch.isalnum() for ch in text):
            continue  # не список или колонка одних маркеров (её pdf2docx делает отдельной ячейкой)
        starts = [run for run in p.findall(qn("w:r")) if (run.findtext(qn("w:t")) or "").startswith("•")]
        if len(starts) < 2:
            continue
        items, started = [p], False
        for run in p.findall(qn("w:r")):
            t = run.find(qn("w:t"))
            if run in starts:
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
        if tbl.tblGrid is None:
            continue
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
        if fonts is None:
            continue
        key = (fonts.get(qn("w:ascii")) or "").replace(" ", "").lower()
        family = "symbol" if key in ("symbol", "opensymbol") else "wingdings" if key.startswith("wingdings") else ""
        texts = [t for t in run.iter(qn("w:t")) if t.text]
        bullets = SYMBOL_BULLETS.get(family, {})
        if family and texts and all(ch in bullets or ch.isspace() for t in texts for ch in t.text):
            for t in texts:  # маркер списка «квадратиком» -> обычный символ
                t.text = "".join(bullets.get(ch, ch) for ch in t.text)
            new_name = "Arial"
        else:
            new_name = next((name for prefix, name in FONT_NAMES.items() if key.startswith(prefix)), None)
        if new_name:
            for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                if fonts.get(qn(attr)) is not None:
                    fonts.set(qn(attr), new_name)
    split_bullets(body)


def docx_chars(element):
    """Буквы и цифры в XML-фрагменте DOCX."""
    return sum(ch.isalnum() for t in element.iter(qn("w:t")) for ch in (t.text or ""))


# ----------------------------------------------------------------------------- конвертация

def open_pdf(pdf_path):
    """Открывает PDF (или картинку — как одностраничный скан) и расшифровывает, если нужен пароль."""
    try:
        src = pymupdf.open(pdf_path)
    except Exception as e:  # noqa: BLE001 — пустой, обрезанный или не тот файл
        raise UserError(f"Не удалось открыть файл (он пустой, повреждён или это не PDF): {e}") from e
    if not src.is_pdf:
        if Path(pdf_path).suffix.lower() not in IMAGE_SUFFIXES:
            raise UserError("Это не PDF-файл и не картинка — конвертирую только PDF (и сканы-картинки).")
        try:
            src = pymupdf.open("pdf", src.convert_to_pdf())
        except Exception as e:  # noqa: BLE001
            raise UserError(f"Не удалось прочитать картинку: {e}") from e
    if src.needs_pass:
        if NO_CONSOLE or sys.stdin is None:
            raise UserError("PDF защищён паролем, а спросить пароль негде (нет консоли).")
        for attempt in range(3):
            try:
                password = getpass.getpass("PDF защищён паролем. Пароль: ")
            except EOFError:
                raise UserError("PDF защищён паролем, а пароль не введён.") from None
            if src.authenticate(password):
                break
            print("Неверный пароль." + (" Попробуйте ещё раз." if attempt < 2 else ""))
        else:
            raise UserError("Неверный пароль.")
    if len(src) == 0:
        raise UserError("В PDF нет ни одной страницы — файл повреждён или скачан не полностью.")
    return src


def parse_pages(pdf_bytes, settings, label_offset=0):
    """Разбор вёрстки pdf2docx. Возвращает (страницы, номера неудачных). Если общий анализ документа
    падает (повреждённый объект на одной странице), страницы разбираются по одной."""
    try:
        cv = Converter(stream=pdf_bytes)
        cv.load_pages().parse_document(**settings)
        pages = list(cv.pages)
    except Exception:  # noqa: BLE001
        pages = None
    if pages is None:
        whole, pages = pymupdf.open("pdf", pdf_bytes), []
        for i in range(len(whole)):
            one = pymupdf.open()
            one.insert_pdf(whole, from_page=i, to_page=i)
            try:
                single = Converter(stream=one.tobytes())
                single.load_pages().parse_document(**settings)
                pages.append(single.pages[0])
            except Exception:  # noqa: BLE001
                pages.append(None)
    failed = []
    for i, page in enumerate(pages):
        if page is None:
            failed.append(i)
            continue
        try:
            page.parse(**settings)
        except Exception as e:  # noqa: BLE001
            warn(f"Стр. {i + 1 + label_offset}: ошибка разбора вёрстки ({e}), распознаю через OCR")
            failed.append(i)
    return pages, failed


def page_lost_text(page, expected):
    """Пробная сборка страницы: pdf2docx иногда молча теряет текст (вертикальный, повёрнутый) или падает."""
    try:
        probe = Document()
        page.make_docx(probe)
    except Exception:  # noqa: BLE001
        return True
    return expected >= MIN_CHARS and docx_chars(probe.element.body) < 0.5 * expected


def convert(pdf_path, out=None, lang="rus+eng", dpi=300, force_ocr=False):
    """Конвертирует один PDF (или картинку); out — файл .docx или папка (по умолчанию рядом с PDF).
    Возвращает путь сохранённого DOCX."""
    src = open_pdf(pdf_path)
    docx_path, overwrite = output_path(pdf_path, out)
    docx_path, overwrite = writable_target(docx_path, overwrite)
    n = len(src)
    scans = [i for i in range(n) if force_ocr or not has_text_layer(src[i])]
    ocr_lang, ocr_error = None, None
    if scans:
        try:
            ocr_lang = setup_tesseract(lang)
        except OcrUnavailable as e:
            ocr_error = str(e)
            warn(ocr_error)
    print(f"Страниц: {n}, с текстом: {n - len(scans)}, сканов (OCR): {len(scans)}")

    bar = tqdm(total=(len(scans) if ocr_lang else 0) * COST_OCR + n * (COST_ANALYZE + COST_PARSE + COST_BUILD),
               desc="Конвертация", bar_format="{desc}: {percentage:3.0f}%|{bar}| {elapsed}<{remaining}{postfix}",
               file=sys.stderr, disable=sys.stderr is None)

    def progress(done, total):
        bar.set_postfix_str(f"распознавание текста: {done}/{total} стр.")
        bar.update(COST_OCR)

    # 1. Копия документа (со всеми слоями и без шифрования), в которой сканы заменяются «чистыми» страницами
    work = pymupdf.open("pdf", src.tobytes(encryption=pymupdf.PDF_ENCRYPT_NONE))
    if scans and ocr_lang:
        bar.set_postfix_str(f"распознавание текста: 0/{len(scans)} стр.")
        done = [0]

        def replace_page(i, ocr, size_dpi):
            done[0] += 1
            if isinstance(ocr, Exception):
                warn(f"Стр. {i + 1}: не удалось распознать ({str(ocr).strip()[:200]}) — оставлена картинкой")
            elif ocr["chars"] >= MIN_CHARS or force_ocr:  # иначе текста нет (например, фото) — как есть
                draw_ocr_page(work, i, ocr, size_dpi)
                work.delete_page(i + 1)
            progress(done[0], len(scans))

        recognize_pages(src, scans, ocr_lang, dpi, replace_page)

    # 2. Разбор вёрстки всех страниц через pdf2docx (неудачные — повторно через OCR, затем картинкой)
    settings = Converter.default_settings.fget(None)  # свойство pdf2docx, экземпляр ему не нужен
    bar.set_postfix_str("анализ структуры документа")
    pages, failed = parse_pages(work.tobytes(garbage=1), settings)
    bar.update(n * COST_ANALYZE)
    replaced = set(scans) if ocr_lang else set()
    for i, page in enumerate(pages):
        bar.set_postfix_str(f"стр. {i + 1}/{n}: разбор вёрстки")
        if page is not None and i not in failed:
            expected = sum(ch.isalnum() for ch in visible_text(work[i]))
            if page_lost_text(page, expected):
                warn(f"Стр. {i + 1}: текст не удалось перенести как есть, распознаю через OCR")
                failed.append(i)
        bar.update(COST_PARSE)

    retry_ocr = [i for i in failed if i not in replaced]  # страницы, ещё не распознававшиеся через OCR
    if retry_ocr and ocr_lang is None and ocr_error is None:
        try:
            ocr_lang = setup_tesseract(lang)
        except OcrUnavailable as e:
            ocr_error = str(e)
            warn(ocr_error)
    if failed:
        retry, recognized = pymupdf.open(), {}
        if retry_ocr and ocr_lang:
            bar.total += len(retry_ocr) * COST_OCR
            done = [0]

            def keep(i, ocr, size_dpi):
                done[0] += 1
                if not isinstance(ocr, Exception) and ocr["chars"] >= MIN_CHARS:
                    recognized[i] = (ocr, size_dpi)
                progress(done[0], len(retry_ocr))

            recognize_pages(src, retry_ocr, ocr_lang, dpi, keep)
        order = sorted(failed)
        for i in order:
            if i in recognized:
                draw_ocr_page(retry, -1, *recognized[i])
            else:
                image_page(retry, src[i])
        retry_pages, still_failed = parse_pages(retry.tobytes(garbage=1), settings)
        for j, i in enumerate(order):
            if j not in still_failed and retry_pages[j] is not None:
                pages[i] = retry_pages[j]
                if i not in recognized:
                    warn(f"Стр. {i + 1}: вставлена в документ картинкой (текст не распознан)")
            else:
                pages[i] = None
                warn(f"Стр. {i + 1}: не удалось перенести страницу")

    # 3. Сборка DOCX
    doc = Document()
    body = doc.element.body
    for i, page in enumerate(pages):
        bar.set_postfix_str(f"стр. {i + 1}/{n}: сборка документа")
        if page is not None and page.finalized:
            before = len(body)
            try:
                page.make_docx(doc)
            except Exception as e:  # noqa: BLE001 — убираем недособранное, чтобы не было «полстраницы»
                for element in list(body)[before - 1:-1]:
                    body.remove(element)
                warn(f"Стр. {i + 1}: не удалось собрать страницу ({e})")
        bar.update(COST_BUILD)
    polish_docx(doc)
    bar.set_postfix_str("сохранение")
    saved = save_docx(doc, docx_path, overwrite)
    bar.set_postfix_str("готово")
    bar.close()
    work.close()
    src.close()
    return saved



# ----------------------------------------------------------------------------- файлы и запуск

NO_CONSOLE = False  # программа собрана без консоли (--noconsole): итог показывается окном


def output_path(pdf_path, out):
    """(путь DOCX, можно ли его перезаписать): рядом с PDF, в папку out или в файл out (только явно указанный
    файл перезаписывается). Результат никогда не совпадает с исходным файлом и всегда имеет расширение .docx."""
    pdf = Path(pdf_path)
    if out and Path(out).is_dir():
        target, explicit = Path(out) / (pdf.stem + ".docx"), False
    elif out:
        target, explicit = Path(out), True
    else:
        target, explicit = pdf.with_suffix(".docx"), False
    if target.suffix.lower() != ".docx":
        raise UserError(f"Результат должен быть файлом .docx, а не «{target.name}».")
    if os.path.normcase(os.path.abspath(target)) == os.path.normcase(os.path.abspath(pdf)):
        raise UserError("Результат совпадает с исходным файлом.")
    return target, explicit


def can_write(folder):
    try:
        with tempfile.TemporaryFile(dir=folder):
            return True
    except OSError:
        return False


def writable_target(path, overwrite):
    """До долгой конвертации проверяет, что в папку можно записать (диск только для чтения, сетевая папка,
    «Контролируемый доступ к папкам» Защитника Windows); иначе сохраняем в «Документы» или домашнюю папку."""
    folder = path.parent
    if not folder.is_dir():
        raise UserError(f"Папка для результата не существует: {folder}")
    if can_write(folder):
        return path, overwrite
    for fallback in (Path.home() / "Documents", Path.home(), Path(tempfile.gettempdir())):
        if fallback.is_dir() and can_write(fallback):
            warn(f"В папку {folder} нельзя записать — результат будет сохранён в {fallback}")
            return fallback / path.name, False
    raise UserError(f"Нет доступа на запись в папку {folder}")


def save_docx(doc, path, overwrite=False):
    """Сохраняет DOCX. Существующие файлы не трогает: если имя занято (или файл открыт в Word),
    берётся «файл (1).docx», «файл (2).docx»... Перезаписывается только явно указанный файл."""
    if overwrite:
        try:
            doc.save(str(path))
            return path
        except PermissionError:
            pass  # файл открыт в Word — сохраним рядом под свободным именем
    for k in range(1 if overwrite else 0, 1000):
        target = path if k == 0 else path.with_name(f"{path.stem} ({k}){path.suffix}")
        try:
            with open(target, "xb") as file:
                try:
                    doc.save(file)
                except BaseException:
                    file.close()
                    target.unlink(missing_ok=True)
                    raise
            return target
        except FileExistsError:
            continue
        except PermissionError as e:
            raise UserError(f"Нет доступа на запись: {target}") from e
    raise UserError(f"Не удалось подобрать свободное имя для {path}")


def open_file(path):
    """Открывает файл в программе по умолчанию (Word, LibreOffice, ...)."""
    try:
        if sys.platform == "win32":
            os.startfile(str(path))
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except OSError as e:
        print(f"Не удалось открыть файл автоматически: {e}")


def pick_pdf_windows():
    """Стандартное окно Windows «Открыть файл» с выбором нескольких файлов (comdlg32, без tkinter)."""
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

    images = ";".join(f"*{suffix}" for suffix in sorted(IMAGE_SUFFIXES))
    filters = wide(f"PDF и сканы-картинки\0*.pdf;{images}\0PDF (*.pdf)\0*.pdf\0Все файлы\0*.*\0\0")
    title = wide("Выберите PDF для конвертации в Word (можно несколько)\0")
    buffer = ctypes.create_unicode_buffer(65536)
    ofn = OPENFILENAMEW(lStructSize=ctypes.sizeof(OPENFILENAMEW),
                        hwndOwner=ctypes.windll.kernel32.GetConsoleWindow(),
                        lpstrFilter=ctypes.addressof(filters), lpstrFile=ctypes.addressof(buffer),
                        nMaxFile=len(buffer), lpstrTitle=ctypes.addressof(title),
                        Flags=0x00081A08)  # EXPLORER | FILEMUSTEXIST | PATHMUSTEXIST | ALLOWMULTISELECT | NOCHANGEDIR
    if not ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):
        return []
    parts = buffer[:].split("\0")
    parts = parts[:parts.index("")] if "" in parts else parts
    if len(parts) == 1:
        return parts  # один файл — полный путь
    return [os.path.join(parts[0], name) for name in parts[1:]]  # папка, затем имена файлов


def pick_pdf():
    """Окно выбора PDF (если файлы не переданы в командной строке). Возвращает список путей."""
    try:
        if sys.platform == "win32":
            return pick_pdf_windows()
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        paths = filedialog.askopenfilenames(title="Выберите PDF",
                                            filetypes=[("PDF", "*.pdf"), ("Все файлы", "*.*")])
        root.destroy()
        return list(paths)
    except Exception:  # noqa: BLE001 — нет графики: спросим путь в консоли
        if NO_CONSOLE:
            return []
        try:
            path = input("Путь к PDF-файлу: ").strip().strip('"')
        except EOFError:
            return []
        return [path] if path else []


def setup_streams():
    """Вывод в файл или канал — в UTF-8 (в Windows иначе cp1252 и кириллица не пишется); если консоли нет
    (сборка --noconsole), вывод уходит в никуда, а итог показывается окном."""
    global NO_CONSOLE
    for name in ("stdout", "stderr"):
        stream = getattr(sys, name)
        if stream is None:
            NO_CONSOLE = True
            setattr(sys, name, open(os.devnull, "w", encoding="utf-8"))
        elif not stream.isatty() and hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def own_console():
    """Окно консоли создано Windows только для этой программы (двойной щелчок, перетаскивание файлов на
    значок, «Открыть с помощью») и закроется сразу после неё — тогда перед выходом нужна пауза."""
    if sys.platform != "win32" or NO_CONSOLE:
        return False
    try:
        import ctypes
        pids = (ctypes.c_uint32 * 16)()
        count = ctypes.windll.kernel32.GetConsoleProcessList(pids, 16)
    except (OSError, AttributeError):
        return False
    meipass = getattr(sys, "_MEIPASS", None)  # у однофайлового exe к консоли подключён ещё и загрузчик
    onefile = bool(meipass) and Path(meipass).resolve().parent != Path(sys.executable).resolve().parent
    return 0 < count <= (2 if onefile else 1)


def message_box(text, error=False):
    """Итог для сборки без консоли."""
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, text[-3000:], "PDF2Word", 0x10 if error else 0x40)
        except (OSError, AttributeError):
            pass


def dpi_value(text):
    try:
        value = int(text)
    except ValueError:
        raise argparse.ArgumentTypeError("нужно целое число") from None
    if not 100 <= value <= 600:
        raise argparse.ArgumentTypeError("допустимо от 100 до 600")
    return value


def main():
    ap = argparse.ArgumentParser(prog="PDF2Word", description="Конвертер PDF -> DOCX с распознаванием сканов")
    ap.add_argument("files", nargs="*", metavar="PDF",
                    help="PDF-файлы или сканы-картинки (если не указаны — откроется окно выбора)")
    ap.add_argument("-o", "--output", help="куда сохранить: файл .docx (для одного PDF) или папка")
    ap.add_argument("--lang", default="rus+eng", help="языки OCR через + (по умолчанию rus+eng)")
    ap.add_argument("--dpi", type=dpi_value, default=300, help="разрешение для OCR, 100-600 (по умолчанию 300)")
    ap.add_argument("--ocr", action="store_true", help="распознавать через OCR все страницы")
    ap.add_argument("--no-open", action="store_true", help="не открывать результат после конвертации")
    args = ap.parse_args()

    files, out = [], args.output
    for name in args.files:  # в Windows маски («*.pdf») раскрывает не консоль, а сама программа
        matches = sorted(glob.glob(name)) if any(ch in name for ch in "*?[") and not os.path.exists(name) else []
        files += matches or [name]
    if len(files) == 2 and not out and Path(files[1]).suffix.lower() == ".docx" \
            and Path(files[0]).suffix.lower() != ".docx" and not Path(files[1]).exists():
        files, out = files[:1], files[1]  # прежняя форма вызова: «pdf2word.py вход.pdf выход.docx»
    if not files:
        files = pick_pdf()
        if not files:
            print("Файл не выбран.")
            return 0
    files = list({os.path.normcase(os.path.abspath(f)): f for f in files}.values())  # без повторов
    if out and len(files) > 1 and not Path(out).is_dir():
        raise UserError("Для нескольких файлов в -o нужно указать существующую папку.")

    saved, failed = [], []
    for number, pdf_path in enumerate(files, 1):
        if len(files) > 1:
            print(f"\n[{number}/{len(files)}] {pdf_path}")
        try:
            if not os.path.isfile(pdf_path):
                raise UserError(f"Файл не найден: {pdf_path}")
            result = convert(pdf_path, out, args.lang, args.dpi, args.ocr)
        except UserError as e:
            print(f"Ошибка: {e}")
            failed.append(pdf_path)
            continue
        except Exception:  # noqa: BLE001 — неожиданная ошибка одного файла не мешает остальным
            print(f"Ошибка при конвертации {pdf_path}:")
            traceback.print_exc()
            failed.append(pdf_path)
            continue
        print(f"Готово: {result}")
        saved.append(result)
    if len(files) > 1:
        print(f"\nСконвертировано: {len(saved)} из {len(files)}")
    if saved and not args.no_open:
        if len(saved) <= 5:
            for result in saved:
                open_file(result)
        else:  # много файлов — открываем папку, а не 20 окон Word
            open_file(saved[0].parent)
    return 1 if failed else 0


if __name__ == "__main__":
    setup_streams()
    pause = own_console() or len(sys.argv) == 1  # окно закроется само — дадим прочитать ошибки
    code = 0
    try:
        code = main()
    except KeyboardInterrupt:
        stop_tesseract()
        print("\nОтменено.")
        code = 130
    except UserError as e:
        print(f"Ошибка: {e}")
        code = 1
    except SystemExit as e:  # argparse: --help или неверные параметры
        code = e.code if isinstance(e.code, int) else (0 if e.code is None else 1)
        if isinstance(e.code, str):
            print(e.code)
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        code = 1
    if NO_CONSOLE and (code or WARNINGS):
        message_box("Есть ошибки или предупреждения:\n\n" + "\n".join(WARNINGS) if WARNINGS
                    else "Не удалось сконвертировать файл.", error=bool(code))
    elif pause and (code or WARNINGS):
        try:
            input("\nНажмите Enter, чтобы закрыть окно...")
        except (EOFError, OSError):
            pass
    sys.exit(code)
