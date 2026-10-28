"""Пересчитать книгу и снять PNG только указанных листов.

    python tools/render.py файл.xlsx --sheets "Свод,Работы,Материалы" --out outbox/<лот>/png

Остальные листы не рендерятся. Если LibreOffice не установлен — один раз
bash tools/bootstrap.sh
"""
import argparse
import shutil
import subprocess
import tempfile
from pathlib import Path

import openpyxl


def main():
    p = argparse.ArgumentParser()
    p.add_argument("xlsx")
    p.add_argument("--sheets", default="Свод,Работы,Материалы")
    p.add_argument("--out", default="png")
    a = p.parse_args()
    if not shutil.which("soffice"):
        raise SystemExit("нет LibreOffice: bash tools/bootstrap.sh")
    wanted = [s.strip() for s in a.sheets.split(",") if s.strip()]
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(a.xlsx)
    for ws in wb.worksheets:
        ws.sheet_state = "visible" if ws.title in wanted else "hidden"
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "book.xlsx"
        wb.save(src)
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, str(src)],
                       check=True, capture_output=True)
        pdf = Path(tmp) / "book.pdf"
        subprocess.run(["pdftoppm", "-png", "-r", "100", str(pdf), str(out / "sheet")],
                       check=True, capture_output=True)
    print(f"PNG: {out} ({len(list(out.glob('*.png')))} стр., листы: {', '.join(wanted)})")


if __name__ == "__main__":
    main()
