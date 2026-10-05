#!/usr/bin/env python3
r"""
Сборка PDF2Word.exe для Windows: один файл со встроенным Tesseract OCR и иконкой.

    pip install -r requirements.txt pyinstaller
    python build_exe.py                        # Tesseract из C:\Program Files\Tesseract-OCR
    python build_exe.py "D:\Tesseract-OCR"     # или из своей папки

Результат: dist\PDF2Word.exe — на другом компьютере ничего устанавливать не нужно.
"""
import shutil
import sys
from pathlib import Path

import pefile
import PyInstaller.__main__

ROOT = Path(__file__).resolve().parent
BUILD = ROOT / "build"
LANGS = ["rus", "eng", "osd"]  # osd нужен для автоповорота страниц

SPEC = """
a = Analysis([{script!r}], datas=[({tesseract!r}, "tesseract"), ({fonts!r}, "assets/fonts")],
             excludes=["tkinter"])  # в Windows окно выбора файла системное, tkinter не нужен
# видеомодуль OpenCV (ffmpeg, ~30 МБ) конвертеру не нужен
a.binaries = [b for b in a.binaries if "opencv_videoio_ffmpeg" not in b[0]]
exe = EXE(PYZ(a.pure), a.scripts, a.binaries, a.datas, name="PDF2Word",
          icon={icon!r}, console=True, upx=False)
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
    tesseract_src = Path(sys.argv[1] if len(sys.argv) > 1 else r"C:\Program Files\Tesseract-OCR")
    tesseract_dst = BUILD / "tesseract"
    copy_tesseract(tesseract_src, tesseract_dst)

    spec = BUILD / "PDF2Word.spec"
    spec.write_text(SPEC.format(script=str(ROOT / "pdf2word.py"), tesseract=str(tesseract_dst),
                                fonts=str(ROOT / "assets" / "fonts"), icon=str(ROOT / "assets" / "icon.ico")),
                    encoding="utf-8")
    PyInstaller.__main__.run([str(spec), "--distpath", str(ROOT / "dist"),
                              "--workpath", str(BUILD / "pyinstaller"), "--noconfirm", "--clean"])
    print(f"\nГотово: {ROOT / 'dist' / 'PDF2Word.exe'}")


if __name__ == "__main__":
    main()
