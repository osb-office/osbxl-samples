# rows-cols.py — node 版サンプルと同一検証 (pyosbxl)
import re
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("rows-cols-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# 行 → 列 の順に並べて 1 行で表示する（· は空セル）
def dump(label, rng):
    v = ws.getRange(rng).getValue()
    def keyf(k):
        m = re.match(r"^([A-Z]+)(\d+)$", k)
        return (int(m.group(2)), app.convertToColumnNumber(m.group(1)))
    keys = sorted(v.keys(), key=keyf)
    print(label, " ".join(f"{k}={v[k] or '·'}" for k in keys))

# A1:C5 を r{行}c{列} で埋める
def fill():
    for r in range(1, 6):
        for c in range(1, 4):
            ws.getCells(r, c).setValue(f"r{r}c{c}")

fill()
dump("[初期]       A列:", "A1:A5")

# ---------- 行 ----------
ws.insertRow(3, 2)                     # 3 行目の「前」に 2 行挿入 → 既存データは下へずれる
dump("[insertRow]  A列:", "A1:A7")
ws.deleteRow(3, 2)                     # 3 行目から 2 行削除 → 下の行が上がる
dump("[deleteRow]  A列:", "A1:A5")
ws.clearRow(2, 1)                      # 2 行目をクリア（位置はずらさない）
dump("[clearRow]   A列:", "A1:A5")
ws.copyRow("Sheet1", 1, 1, 4)          # 自シートの 1 行目を 4 行目へ上書きコピー
dump("[copyRow]    A列:", "A1:A5")

fill()
# ---------- 列 ----------
ws.insertCol(2, 1)                     # B 列の前に 1 列挿入
dump("[insertCol] 1行目:", "A1:D1")
ws.deleteCol(2, 1)                     # 挿入した B 列を削除
dump("[deleteCol] 1行目:", "A1:C1")
ws.clearCol(3, 1)                      # C 列をクリア
dump("[clearCol]  1行目:", "A1:C1")
fill()
ws.copyCol("Sheet1", 1, 2, 5)          # A,B 列を E,F 列へコピー
dump("[copyCol]   1行目:", "A1:F1")

# ---------- セル範囲 ----------
ws.insertCell("B2:B3", pyosbxl.enums.XlInsertDirection.InsertDirectionDown)
dump("[insertCell] B列:", "B1:B5")
ws.deleteCell("B2:B3", pyosbxl.enums.XlInsertDirection.DeleteDirectionUp)
dump("[deleteCell] B列:", "B1:B5")
ws.clearCell("A1:B2")
dump("[clearCell] A1:B2:", "A1:B2")
ws.copyCell("Sheet1", "A3:B4", "F8")
dump("[copyCell]  F8:G9:", "F8:G9")

wb.save()
wb.close()