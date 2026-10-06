# -*- mode: python ; coding: utf-8 -*-
r"""
Сборка PDF2Word для Windows:

    pip install -r requirements.txt pyinstaller
    pyinstaller PDF2Word.spec

Результат — папка dist\PDF2Word с PDF2Word.exe. Один exe-файл вместо папки (запускается дольше):
    set PDF2WORD_ONEFILE=1
    pyinstaller PDF2Word.spec

Tesseract берётся из установленного (C:\Program Files\Tesseract-OCR или папка из переменной
TESSERACT_DIR): в сборку попадают только tesseract.exe, нужные ему DLL, языки rus/eng/osd и конфиги.
Если Tesseract не найден, программа соберётся без него и будет искать установленный Tesseract при запуске.
"""
import os
from pathlib import Path

ROOT = Path(SPECPATH)
ONEFILE = os.environ.get("PDF2WORD_ONEFILE", "").strip() not in ("", "0")
LANGS = ["rus", "eng", "osd"]  # osd нужен для автоповорота страниц
TESSERACT_DIRS = [
    os.environ.get("TESSERACT_DIR", ""),
    r"C:\Program Files\Tesseract-OCR",
    r"C:\Program Files (x86)\Tesseract-OCR",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR"),
    os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR"),
]


def say(ru, en):
    """Сообщение сборки; если консоль не умеет кириллицу (вывод перенаправлен), — по-английски."""
    try:
        print(ru)
    except UnicodeEncodeError:
        print(en)


def tesseract_files():
    """(файл, папка в сборке): tesseract.exe, только нужные ему DLL, языки и конфиги для hocr/tsv."""
    src = next((Path(d) for d in TESSERACT_DIRS if d and (Path(d) / "tesseract.exe").is_file()), None)
    if src is None:
        say("\nВНИМАНИЕ: Tesseract не найден (укажите папку в TESSERACT_DIR). Программа соберётся без него:"
            "\nсканы будут распознаваться, только если Tesseract установлен на компьютере.\n",
            "\nWARNING: Tesseract not found (set TESSERACT_DIR). Building without it: scanned pages will be"
            "\nrecognized only if Tesseract is installed on the computer.\n")
        return []
    import pefile  # ставится вместе с PyInstaller на Windows

    local = {f.name.lower(): f for f in src.glob("*.dll")}
    files, todo, seen = [], [src / "tesseract.exe"], set()
    while todo:  # зависимости: tesseract.exe -> libtesseract -> leptonica -> ...
        path = todo.pop()
        files.append((str(path), "tesseract"))
        pe = pefile.PE(str(path), fast_load=True)
        pe.parse_data_directories(directories=[pefile.DIRECTORY_ENTRY["IMAGE_DIRECTORY_ENTRY_IMPORT"],
                                               pefile.DIRECTORY_ENTRY["IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT"]])
        for imp in getattr(pe, "DIRECTORY_ENTRY_IMPORT", []) + getattr(pe, "DIRECTORY_ENTRY_DELAY_IMPORT", []):
            name = imp.dll.decode(errors="replace").lower()
            if name in local and name not in seen:
                seen.add(name)
                todo.append(local[name])
        pe.close()

    tessdata = src / "tessdata"
    for lang in LANGS:
        data = tessdata / f"{lang}.traineddata"
        if not data.is_file():
            raise SystemExit(f"Missing {data}: reinstall Tesseract with the Russian language "
                             "(Additional language data -> Russian).")
        files.append((str(data), "tesseract/tessdata"))
    for config in (tessdata / "configs").glob("*"):  # hocr, tsv и др. — нужны для вывода Tesseract
        files.append((str(config), "tesseract/tessdata/configs"))
    say(f"\nTesseract: {src} (файлов программы: {len(seen) + 1})\n", f"\nTesseract: {src} ({len(seen) + 1} program files)\n")
    return files


def runtime_dlls(collected):
    """Библиотеки Visual C++ (их требует PyMuPDF): если PyInstaller их не собрал, берём из System32 —
    тогда программа запустится и там, где Visual C++ Redistributable не установлен."""
    system32 = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32"
    names = {Path(dest).name.lower() for dest, *_ in collected}
    extra = []
    for name in ("msvcp140.dll", "msvcp140_1.dll", "msvcp140_2.dll", "vcruntime140.dll", "vcruntime140_1.dll"):
        path = system32 / name
        if name not in names and path.is_file() and b"Wine " not in path.read_bytes()[:0x100]:  # не заглушка Wine
            extra.append((name, str(path), "BINARY"))
    return extra


fonts = [(str(f), "assets/fonts") for f in (ROOT / "assets" / "fonts").glob("*")]
if not fonts:
    raise SystemExit("Missing assets/fonts (LiberationSerif-*.ttf): they are required for recognized text.")
a = Analysis([str(ROOT / "pdf2word.py")], datas=fonts,
             excludes=["tkinter", "PIL"])  # в Windows окно выбора файла системное; Pillow не используется
# видеомодуль OpenCV (ffmpeg, ~30 МБ) конвертеру не нужен
a.binaries = [b for b in a.binaries if "opencv_videoio_ffmpeg" not in b[0]]
a.binaries += runtime_dlls(a.binaries)
# Tesseract кладётся как есть, в обход анализа PyInstaller (иначе его DLL попадут в сборку дважды)
tesseract = [(f"{folder}/{Path(path).name}", path, "DATA") for path, folder in tesseract_files()]
pyz = PYZ(a.pure)
icon = str(ROOT / "assets" / "icon.ico")
if ONEFILE:
    exe = EXE(pyz, a.scripts, a.binaries, a.datas, tesseract, name="PDF2Word", icon=icon, console=True, upx=False)
else:
    exe = EXE(pyz, a.scripts, exclude_binaries=True, name="PDF2Word", icon=icon, console=True, upx=False)
    coll = COLLECT(exe, a.binaries, a.datas, tesseract, name="PDF2Word", upx=False)
