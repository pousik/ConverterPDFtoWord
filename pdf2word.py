#!/usr/bin/env python3
"""
PDF -> Word (DOCX): простой конвертер с распознаванием сканов (OCR) и прогрессом.

    python pdf2word.py документ.pdf              -> документ.docx рядом с PDF, сразу открывается
    python pdf2word.py документ.pdf итог.docx
    python pdf2word.py                           -> откроется окно выбора PDF

Страницы с текстовым слоем конвертирует pdf2docx (сохраняются шрифты, таблицы,
картинки). Страницы-сканы распознаёт Tesseract OCR и превращает в обычный
редактируемый текст.
"""
import argparse
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

import numpy as np
import pymupdf
import pytesseract
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt
from PIL import ImageFilter
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

OCR_FONT = "Times New Roman"
MIN_CHARS = 20          # меньше букв/цифр на странице -> считаем её сканом
SIZE_FACTOR = 1.08      # x_size строки от Tesseract -> кегль шрифта (подобрано на тестах)
LINE_HEIGHT = 1.15      # одинарный интервал в Word ~ 1.15 кегля
TAB_GAP = 1.2           # промежуток между словами больше 1.2 кегля -> колонка таблицы (табуляция)
XHTML = "{http://www.w3.org/1999/xhtml}"
# «Вес» шагов в прогресс-баре, чтобы проценты шли примерно равномерно по времени
COST_OCR, COST_ANALYZE, COST_PARSE, COST_BUILD = 25, 1, 2, 1

# Tesseract, встроенный в PDF2Word.exe (см. build_exe.py)
BUNDLED_TESSERACT = Path(getattr(sys, "_MEIPASS", Path(__file__).parent)) / "tesseract"
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


def ocr_page(page, lang, dpi):
    """Распознаёт страницу. Возвращает (ширина, высота, абзацы); абзац — список строк, всё в пунктах."""
    img = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY).pil_image()
    img = img.filter(ImageFilter.MedianFilter(3))  # убирает «пыль» скана, заметно повышает точность
    try:  # автоповорот страниц, отсканированных боком или вверх ногами
        angle = pytesseract.image_to_osd(img, output_type=pytesseract.Output.DICT)["rotate"]
        if angle:
            img = img.rotate(-angle, expand=True)
    except pytesseract.TesseractError:
        pass  # слишком мало текста для определения ориентации

    hocr = pytesseract.image_to_pdf_or_hocr(img, lang=lang, config="--oem 1 --psm 3", extension="hocr")
    ink = np.asarray(img) < 128
    k = 72 / dpi  # пиксели -> пункты
    pars = []
    for par in ET.fromstring(hocr).iter(f"{XHTML}p"):
        lines = []
        for line in par:
            words = []
            for w in line:
                text = "".join(w.itertext()).strip()
                conf = (hocr_value(w, "x_wconf") or [100])[0]
                if text and (conf >= 30 or any(ch.isalnum() for ch in text)):  # иначе это грязь скана
                    x0, _, x1, _ = hocr_value(w, "bbox")
                    words.append((x0 * k, x1 * k, text))
            if not words:
                continue
            box = hocr_value(line, "bbox")
            x0, y0, x1, y1 = (v * k for v in box)
            _, offset = hocr_value(line, "baseline") or (0, 0)
            x_size = max((hocr_value(line, "x_size") or [box[3] - box[1]])[0], 1)
            lines.append({"words": words, "x0": x0, "y0": y0, "x1": x1, "y1": y1,
                          "base": y1 + offset * k, "size": x_size * k * SIZE_FACTOR,
                          "stroke": stroke_width(ink, box) / x_size})
        pars.append(lines)

    # жирный шрифт: штрихи заметно толще, чем у основного текста страницы
    strokes = [l["stroke"] for lines in pars for l in lines]
    normal = statistics.median(strokes) if strokes else 0
    paragraphs = []
    for lines in pars:
        for line in lines:
            line["bold"] = line["stroke"] > normal * 1.3
        paragraphs += split_paragraph(lines)
    return img.width * k, img.height * k, paragraphs


def stroke_width(ink, box):
    """Толщина штриха букв: медиана длин горизонтальных отрезков «чернил» в рамке строки."""
    x0, y0, x1, y1 = (int(v) for v in box)
    edges = np.diff(np.pad(ink[y0:y1, x0:x1], ((0, 0), (1, 1))).astype(np.int8), axis=1)
    runs = np.nonzero(edges == -1)[1] - np.nonzero(edges == 1)[1]
    return float(np.median(runs)) if runs.size else 0.0


def line_text(line):
    """Текст строки; большие пробелы (колонки таблицы) превращаются в табуляцию."""
    words = line["words"]
    text = words[0][2]
    for prev, word in zip(words, words[1:]):
        text += ("\t" if word[0] - prev[1] > line["size"] * TAB_GAP else " ") + word[2]
    return text


def split_paragraph(lines):
    """Делит абзац Tesseract там, где меняется шрифт, большой отступ или строка таблицы."""
    result = []
    for line in lines:
        prev = result[-1][-1] if result else None
        if (prev is None or "\t" in line_text(line) or "\t" in line_text(prev)
                or line["bold"] != prev["bold"]
                or abs(line["size"] - prev["size"]) > 0.15 * prev["size"]
                or line["y0"] - prev["y1"] > prev["size"]):
            result.append([line])
        else:
            result[-1].append(line)
    return result


def join_lines(lines):
    """Склеивает строки абзаца, убирая переносы слов («сло-» + «во» -> «слово»)."""
    text = ""
    for line in lines:
        part = line_text(line)
        if text.endswith("-") and len(text) > 1 and text[-2].isalpha() and part[:1].islower():
            text = text[:-1] + part
        else:
            text = f"{text} {part}" if text else part
    return text


def add_ocr_page(doc, width, height, paragraphs):
    """Добавляет распознанную страницу в документ обычными редактируемыми абзацами."""
    section = doc.add_section(WD_SECTION.NEW_PAGE) if doc.paragraphs else doc.sections[0]
    section.page_width, section.page_height = Pt(width), Pt(height)
    if not paragraphs:
        doc.add_paragraph()
        return

    left = min(l["x0"] for p in paragraphs for l in p)
    right = max(l["x1"] for p in paragraphs for l in p)
    top = min(p[0]["y0"] for p in paragraphs)
    section.left_margin = Pt(min(max(left, 18), width / 3))
    section.right_margin = Pt(min(max(width - right, 18), width / 3))
    section.top_margin = Pt(min(max(top, 18), height / 3))
    section.bottom_margin = Pt(36)
    left, right = section.left_margin.pt, width - section.right_margin.pt
    center = (left + right) / 2
    # выравнивание по ширине: почти все строки (кроме последних в абзацах) доходят до правого края
    inner = [l for p in paragraphs for l in p[:-1]]
    justified = len(inner) > 2 and sum(right - l["x1"] < l["size"] for l in inner) > 0.8 * len(inner)

    prev_base = top
    for lines in paragraphs:
        size = round(statistics.median(l["size"] for l in lines) * 2) / 2
        size = min(max(size, 6), 72)
        first, last = lines[0], lines[-1]
        indent = min(l["x0"] for l in lines) - left
        par = doc.add_paragraph()
        fmt = par.paragraph_format
        fmt.space_after = Pt(0)
        fmt.space_before = Pt(min(max(first["base"] - prev_base - size * LINE_HEIGHT, 0), 200))
        if len(lines) > 1:
            pitch = statistics.median(b["base"] - a["base"] for a, b in zip(lines, lines[1:]))
            fmt.line_spacing = round(min(max(pitch / (size * LINE_HEIGHT), 1.0), 3.0), 2)
            if first["x0"] - lines[1]["x0"] > size:
                fmt.first_line_indent = Pt(first["x0"] - lines[1]["x0"])
            if justified and all(right - l["x1"] < size for l in lines[:-1]):
                par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if all(abs((l["x0"] + l["x1"]) / 2 - center) < size and l["x0"] - left > size * 3 for l in lines):
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif len(lines) == 1 and right - first["x1"] < size and indent > (right - left) / 3:
            par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif indent > size:
            fmt.left_indent = Pt(indent)
        text = join_lines(lines)
        if "\t" in text:  # строка таблицы: ставим позиции табуляции по колонкам
            for prev, word in zip(first["words"], first["words"][1:]):
                if word[0] - prev[1] > first["size"] * TAB_GAP:
                    fmt.tab_stops.add_tab_stop(Pt(word[0] - left))
        run = par.add_run(text)
        run.font.name = OCR_FONT
        run.font.size = Pt(size)
        run.font.bold = sum(l["bold"] for l in lines) > len(lines) / 2
        prev_base = last["base"]


# ----------------------------------------------------------------------------- конвертация

def convert(pdf_path, docx_path, lang="rus+eng", dpi=300, force_ocr=False):
    pdf = pymupdf.open(pdf_path)
    password = None
    if pdf.needs_pass:
        password = getpass.getpass("PDF защищён паролем. Пароль: ")
        if not pdf.authenticate(password):
            sys.exit("Неверный пароль.")
    n = len(pdf)
    scans = [i for i in range(n) if force_ocr or not has_text_layer(pdf[i])]
    layout = [i for i in range(n) if i not in scans]
    if scans:
        lang = setup_tesseract(lang)
    print(f"Страниц: {n}, с текстом: {len(layout)}, сканов (OCR): {len(scans)}")

    bar = tqdm(total=len(scans) * COST_OCR + len(layout) * (COST_ANALYZE + COST_PARSE) + n * COST_BUILD,
               desc="Конвертация", bar_format="{desc}: {percentage:3.0f}%|{bar}| "
               "{elapsed}<{remaining}{postfix}")

    def extend_bar(units):
        bar.total += units
        bar.refresh()

    # 1. OCR сканов
    ocr = {}
    for i in scans:
        bar.set_postfix_str(f"стр. {i + 1}/{n}: распознавание текста")
        width, height, paragraphs = ocr_page(pdf[i], lang, dpi)
        text_len = sum(len(line_text(l)) for p in paragraphs for l in p)
        if text_len < MIN_CHARS and not force_ocr:
            layout.append(i)  # текста нет (например, фото) — сохраним страницу как есть, с картинкой
            extend_bar(COST_ANALYZE + COST_PARSE)
        else:
            ocr[i] = (width, height, paragraphs)
        bar.update(COST_OCR)

    # 2. Разметка текстовых страниц через pdf2docx
    cv = None
    if layout:
        layout.sort()
        cv = Converter(pdf_path, password=password)
        settings = cv.default_settings
        bar.set_postfix_str("анализ структуры документа")
        cv.load_pages(pages=layout).parse_document(**settings)
        bar.update(len(layout) * COST_ANALYZE)
        for i in layout:
            bar.set_postfix_str(f"стр. {i + 1}/{n}: разбор вёрстки")
            try:
                cv.pages[i].parse(**settings)
            except Exception as e:  # pdf2docx не справился — распознаем страницу как скан
                tqdm.write(f"Стр. {i + 1}: ошибка разбора ({e}), распознаю через OCR")
                lang = setup_tesseract(lang)
                extend_bar(COST_OCR)
                ocr[i] = ocr_page(pdf[i], lang, dpi)
                bar.update(COST_OCR)
            bar.update(COST_PARSE)

    # 3. Сборка DOCX
    doc = Document()
    for i in range(n):
        bar.set_postfix_str(f"стр. {i + 1}/{n}: сборка документа")
        if i in ocr:
            add_ocr_page(doc, *ocr[i])
        elif cv and cv.pages[i].finalized:
            try:
                cv.pages[i].make_docx(doc)
            except Exception as e:
                tqdm.write(f"Стр. {i + 1}: не удалось собрать страницу ({e})")
        bar.update(COST_BUILD)
    bar.set_postfix_str("сохранение")
    saved = save_docx(doc, Path(docx_path))
    bar.set_postfix_str("готово")
    bar.close()
    if cv:
        cv.close()
    pdf.close()
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
