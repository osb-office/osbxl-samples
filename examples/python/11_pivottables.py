# pivot-sample.py — ピボットテーブル: node 版サンプルと同一検証 (pyosbxl)
import pyosbxl

enums = pyosbxl.enums

app = pyosbxl.App()
wb = app.createWorkBook("pivot-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# ---------- 1. ソースデータ ----------
# 地域 / 商品 / 数量 の 10 行
headers = ["地域", "商品", "数量"]
src_data = [
    ["東", "りんご", 10], ["東", "みかん", 20], ["西", "りんご", 30],
    ["西", "みかん", 15], ["東", "りんご", 25], ["北", "みかん", 40],
    ["北", "りんご",  5], ["西", "りんご", 10], ["東", "みかん", 30],
    ["北", "りんご", 20],
]
for i, h in enumerate(headers):
    col = chr(65 + i)
    ws.getRange(f"{col}1:{col}1").setValue(h)
for ri, row in enumerate(src_data):
    for ci, v in enumerate(row):
        col = chr(65 + ci)
        r = ri + 2
        cell = ws.getRange(f"{col}{r}:{col}{r}")
        if isinstance(v, int):
            cell.setNumberValue(v)
        else:
            cell.setValue(v)
src_total = sum(row[2] for row in src_data)
print(f"[1] ソースデータ: {len(src_data)} 行, 数量の合計 = {src_total}")

# ---------- 2. ピボットテーブル作成 (行=地域, 値=数量 合計) ----------
pt = ws.getPivotTables().addPivotTable("売上集計", "A1:C11", "E1")

rf = pt.getField("地域")
rf.setFieldType(enums.XlPivotFieldType.PivotFieldTypeRow)
df = pt.getField("数量")
df.setFieldType(enums.XlPivotFieldType.PivotFieldTypeData)
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum)

pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
             [rf], [], [], [df])
pt.refreshData()
pt.refresh()
print("[2] ピボットテーブル作成 (行=地域, 値=数量 合計, レイアウト=テーブル形式)")

# ---------- 3. 集計結果の読み戻しと検算 ----------
# ピボット出力: E1=ヘッダ, E2〜=行ラベル, F2〜=値, 最終行=総計
pivot_rows = {}
for r in range(2, 11):
    label = ws.getRange(f"E{r}:E{r}").getValue().get(f"E{r}")
    val   = ws.getRange(f"F{r}:F{r}").getValue().get(f"F{r}")
    if label is not None and label != "":
        pivot_rows[label] = float(val)
        print(f"    {label} = {val}")
grand_total = pivot_rows["総計"]
# node の Number 表示（205）に合わせるため、整数なら int で表示する
gt = int(grand_total) if grand_total == int(grand_total) else grand_total
print(f"[3] 検算: 総計 {gt} === ソース合計 {src_total} → {'OK' if grand_total == src_total else 'NG'}")

# ---------- 4. 集計方法を平均に変更 ----------
df.setDataFieldSubTotalsMethod(enums.XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage)
pt.setFields(pt.getPivotTableSetting(), enums.XlPivotTableLayoutType.TabularForm,
             [rf], [], [], [df])
pt.refreshData()
pt.refresh()
print("[4] 集計方法を合計→平均に変更")
for r in range(2, 11):
    label = ws.getRange(f"E{r}:E{r}").getValue().get(f"E{r}")
    val   = ws.getRange(f"F{r}:F{r}").getValue().get(f"F{r}")
    if label is not None and label != "":
        print(f"    {label} = {val}")

# ---------- 5. フィールド情報の取得 ----------
fields = pt.getFields()
print(f"[5] getFields: {len(fields)} 個 → " + ", ".join(f.getCaption() for f in fields))
print(f"    getLayoutType = {int(pt.getLayoutType())} (TabularForm={int(enums.XlPivotTableLayoutType.TabularForm)})")

# ---------- 6. ピボットテーブルの削除 ----------
ws.getPivotTables().removePivotTable("売上集計")
try:
    ws.getPivotTables().getPivotTable("売上集計")
    print("[6] 削除失敗")
except Exception:
    print("[6] removePivotTable → 削除確認OK")

wb.save()
wb.close()