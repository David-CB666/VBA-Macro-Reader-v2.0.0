#!/usr/bin/env python3
"""Generate screenshot PNG from an Excel file using Excel COM automation."""
import sys, os, pythoncom, win32com.client
from pathlib import Path

def excel_to_png(xlsx_path, png_path, sheet_name=None):
    """Open xlsx, export used range as PNG screenshot."""
    xlsx_path = str(Path(xlsx_path).resolve())
    png_path = str(Path(png_path).resolve())
    Path(png_path).parent.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        wb = excel.Workbooks.Open(xlsx_path)
        ws = wb.Sheets(sheet_name) if sheet_name else wb.Sheets(1)
        ws.Activate()

        # Get used range
        used = ws.UsedRange
        # Export as image
        used.CopyPicture(Format=2)  # xlBitmap

        # Paste to new chart then export
        chart = wb.Charts.Add()
        chart.Paste()
        chart.Export(png_path, "PNG")
        chart.Delete()
        wb.Close(False)

        print(f"✓ Screenshot saved: {png_path}")
        return True
    except Exception as e:
        print(f"✗ Error: {e}")
        return False
    finally:
        if excel:
            excel.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python excel_screenshot.py <input.xlsx> [output.png]")
        sys.exit(1)
    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else str(Path(inp).with_suffix('.png'))
    excel_to_png(inp, out)
