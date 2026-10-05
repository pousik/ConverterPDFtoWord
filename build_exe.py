#!/usr/bin/env python3
r"""
Сборка PDF2Word.exe для Windows со встроенным Tesseract OCR и иконкой.

    pip install -r requirements.txt pyinstaller
    python build_exe.py                        # Tesseract из C:\Program Files\Tesseract-OCR
    python build_exe.py "D:\Tesseract-OCR"     # или из своей папки
    python build_exe.py --onefile              # один exe-файл (запускается дольше)

Результат: папка dist\PDF2Word с PDF2Word.exe — на другом компьютере ничего устанавливать не нужно.
"""
import argparse
import shutil
import sys
from pathlib import Path

import pefile
import PyInstaller.__main__

ROOT = Path(__file__).resolve().parent
BUILD = ROOT / "build"
LANGS = ["rus", "eng", "osd"]  # osd нужен для автоповорота страниц

SPEC = """
# tkinter не нужен (в Windows окно выбора файла системное), Pillow — тоже
a = Analysis([{script!r}], datas=[({tesseract!r}, "tesseract"), ({fonts!r}, "assets/fonts")],
             excludes=["tkinter", "PIL"])
# видеомодуль OpenCV (ffmpeg, ~30 МБ) конвертеру не нужен
a.binaries = [b for b in a.binaries if "opencv_videoio_ffmpeg" not in b[0]]
pyz = PYZ(a.pure)
"""
ONEFILE = """
exe = EXE(pyz, a.scripts, a.binaries, a.datas, name="PDF2Word", icon={icon!r}, console=True, upx=False)
"""
ONEDIR = """
exe = EXE(pyz, a.scripts, exclude_binaries=True, name="PDF2Word", icon={icon!r}, console=True, upx=False)
coll = COLLECT(exe, a.binaries, a.datas, name="PDF2Word", upx=False)
"""


def copy_tesseract(src, dst):
    """Копирует tesseract.exe, только нужные ему DLL, языки и конфиги (hocr)."""
    exe = src / "tesseract.exe"
    if not exe.exists():
        sys.exit(f"Не найден {exe}. Установите Tesseract: https://github.com/UB-Mannheim/tesseract/wiki")
    shutil.rmtree(dst, ignore_errors=True)
    (dst / "tessdata").mkdir(parents=True)

    local_dlls = {f.name.lower(): f for f in src.glob("*.dll")}
    todo, seen = [exe], set()
    while todo:  # обходим зависимости: tesseract.exe -> libtesseract -> leptonica -> ...
        f = todo.pop()
        shutil.copy2(f, dst / f.name)
        pe = pefile.PE(str(f), fast_load=True)
        pe.parse_data_directories(directories=[pefile.DIRECTORY_ENTRY["IMAGE_DIRECTORY_ENTRY_IMPORT"]])
        for imp in getattr(pe, "DIRECTORY_ENTRY_IMPORT", []):
            name = imp.dll.decode().lower()
            if name in local_dlls and name not in seen:
                seen.add(name)
                todo.append(local_dlls[name])

    for lang in LANGS:
        data = src / "tessdata" / f"{lang}.traineddata"
        if not data.exists():
            sys.exit(f"Нет {data}. Переустановите Tesseract, отметив язык Russian.")
        shutil.copy2(data, dst / "tessdata")
    shutil.copytree(src / "tessdata" / "configs", dst / "tessdata" / "configs")


def main():
    if not sys.stdout.isatty():  # вывод в файл: в Windows иначе cp1252 и кириллица не пишется
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="Сборка PDF2Word.exe")
    ap.add_argument("tesseract", nargs="?", default=r"C:\Program Files\Tesseract-OCR", help="папка Tesseract-OCR")
    ap.add_argument("--onefile", action="store_true",
                    help="один exe-файл (удобно передавать, но каждый запуск дольше на распаковку)")
    args = ap.parse_args()
    tesseract_dst = BUILD / "tesseract"
    copy_tesseract(Path(args.tesseract), tesseract_dst)

    spec = BUILD / "PDF2Word.spec"
    template = SPEC + (ONEFILE if args.onefile else ONEDIR)
    spec.write_text(template.format(script=str(ROOT / "pdf2word.py"), tesseract=str(tesseract_dst),
                                    fonts=str(ROOT / "assets" / "fonts"), icon=str(ROOT / "assets" / "icon.ico")),
                    encoding="utf-8")
    PyInstaller.__main__.run([str(spec), "--distpath", str(ROOT / "dist"),
                              "--workpath", str(BUILD / "pyinstaller"), "--noconfirm", "--clean"])
    result = ROOT / "dist" / ("PDF2Word.exe" if args.onefile else "PDF2Word/PDF2Word.exe")
    print(f"\nГотово: {result}")


if __name__ == "__main__":
    main()
