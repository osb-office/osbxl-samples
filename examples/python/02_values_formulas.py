# values-formulas.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("sales-py.xlsx")   # 新規ブック作成
ws = wb.openWorkSheet("Sheet1")            # 既定のシートを開く

# --- 見出し (1行目) ---
# app.convertFromRowColNumber(行, 列) で (1,1) -> "A1" のセルアドレスが得られる
for i, text in enumerate(["商品", "単価", "個数", "金額"]):
    cell = app.convertFromRowColNumber(1, i + 1)
    ws.getRange(f"{cell}:{cell}").setValue(text)

# --- 明細 (2〜4行目) ---
data = [("りんご", 100, 3), ("みかん", 50, 6), ("ぶどう", 150, 2)]
for i, (name, price, qty) in enumerate(data):
    row = i + 2
    ws.getRange(f"A{row}:A{row}").setValue(name)
    ws.getRange(f"B{row}:B{row}").setNumberValue(price)
    ws.getRange(f"C{row}:C{row}").setNumberValue(qty)

# --- 金額列 (D2:D4) ---
# 行ごとに異なる相対参照を入れたいので setRangeFormulas を使う。
# 数式文字列に先頭の "=" は付けない。
formulas = [[f"B{i + 2}*C{i + 2}"] for i in range(len(data))]
ws.setRangeFormulas("D2", formulas)

# --- 合計 (D5) ---
ws.getRange("C5:C5").setValue("合計")
ws.getRange("D5:D5").setFormula("SUM(D2:D4)")

# --- 読み戻し ---
# getValue(True) は「生値」= 数式の計算結果を返す
values = ws.getRange("A1:D5").getValue(True)
print("金額 D2:D4 =", ", ".join([values["D2"], values["D3"], values["D4"]]))
print("合計:", values["D5"])

# --- 保存して閉じる ---
wb.save()
wb.close()