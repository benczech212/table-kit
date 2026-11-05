"""Recalculates the workbook in real Excel so cached values are stored, scans every sheet for errors, and exports Builder and Sheet to PDF for a visual check."""
import os, sys
import win32com.client as com
HERE = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(HERE, "Hero-Ledger.xlsx")
xl = com.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
wb = xl.Workbooks.Open(path)
try:
    xl.CalculateFull()
    errs = 0; forms = 0
    for ws in wb.Worksheets:
        ur = ws.UsedRange
        vals = ur.Value; fs = ur.Formula
        if not isinstance(vals, tuple): continue
        for r, row in enumerate(vals):
            for c, v in enumerate(row):
                if isinstance(fs[r][c], str) and fs[r][c].startswith("="): forms += 1
                if isinstance(v, int) and v < -2146826200:
                    errs += 1
                    if errs < 25: print("ERR", ws.Name, ur.Cells(r + 1, c + 1).Address, fs[r][c][:150])
    print("formulas", forms, "errors", errs)
    B = wb.Worksheets("Builder"); S = wb.Worksheets("Sheet")
    print("sample:", B.Range("C7").Value, "HP", wb.Worksheets("Calc").Range("B13").Value, "AC", wb.Worksheets("Calc").Range("B56").Value, "issues:", B.Range("P28").Value)
    B.Activate(); B.Range("C5").Select()
    wb.Save()
    scratch = sys.argv[1] if len(sys.argv) > 1 else HERE
    B.PageSetup.Zoom = False; B.PageSetup.FitToPagesWide = 1; B.PageSetup.FitToPagesTall = False; B.PageSetup.Orientation = 2
    S.ExportAsFixedFormat(0, os.path.join(scratch, "sheet.pdf"))
    B.ExportAsFixedFormat(0, os.path.join(scratch, "builder.pdf"))
finally:
    wb.Close(False); xl.Quit()
