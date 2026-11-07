# bulk-input.py — 2次元配列による値と数式の一括高速投入 (pyosbxl)
import json
import time
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("bulk-input.xlsx")
ws = wb.openWorkSheet("Sheet1")

# ---------- 1. 2次元配列による値の一括高速投入 (setRangeData) ----------
data = [
    ["商品名", "単価", "数量", "金額"],
    ["りんご", "100", "3"],
    ["みかん", "50", "6"],
    ["ぶどう", "150", "2"],
]
ws.setRangeData("A1", data)

print("[1] setRangeData 投入後:")
for a1 in ["A1", "B2", "C2", "D1"]:
    print(f"    {a1} = {json.dumps(ws.getRange(a1).getValue()[a1], ensure_ascii=False)}")

# ---------- 2. 2次元配列による数式の一括高速投入 (setRangeFormulas) ----------
formulas = [
    ["B2*C2"],
    ["B3*C3"],
    ["B4*C4"],
]
ws.setRangeFormulas("D2", formulas)

print("[2] setRangeFormulas 投入後:")
print("    数式 D2:D4 =", json.dumps(ws.getRange("D2:D4").getFormula(), ensure_ascii=False))
print("    計算値 D2:D4 =", json.dumps(ws.getRange("D2:D4").getValue(True), ensure_ascii=False))

# ---------- 3. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
ROWS, COLS = 2000, 5

# セルごとに getCells() して投入
ws_cell = wb.addWorkSheet("CellByCell")
t = time.perf_counter()
for r in range(1, ROWS + 1):
    for c in range(1, COLS + 1):
        ws_cell.getCells(r, c).setNumberValue(r * 100 + c)
ms_cell = int((time.perf_counter() - t) * 1000)

# 2次元配列で setRangeData() 投入 (最高速)
ws_bulk = wb.addWorkSheet("Bulk")
t = time.perf_counter()
matrix = [[str(r * 100 + c) for c in range(1, COLS + 1)] for r in range(1, ROWS + 1)]
ms_build = int((time.perf_counter() - t) * 1000)
t = time.perf_counter()
ws_bulk.setRangeData("A1", matrix)
ms_apply = int((time.perf_counter() - t) * 1000)

print(f"[3] {ROWS * COLS} セルの投入速度比較:")
print(f"    セルごと getCells().setNumberValue() : {ms_cell} ms")
print(f"    setRangeData (構築{ms_build} + 投入{ms_apply}) : {ms_build + ms_apply} ms (最速)")

a = ws_cell.getRange("E2000:E2000").getValue(True)["E2000"]
b = ws_bulk.getRange("E2000:E2000").getValue(True)["E2000"]
print(f"[検証] 両方式とも E2000 = {b if a == float(b) else f'不一致 {a} / {b}'}")

wb.save()
wb.close()