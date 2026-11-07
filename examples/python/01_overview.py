# basic-flow.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl

app = pyosbxl.App()                               # [1] App を生成

wb = app.createWorkBook("flow-py.xlsx")           # [2] 新規作成
ws = wb.openWorkSheet("Sheet1")                   # [3] ワークシートを開く

ws.getRange("A1:A1").setValue("新規作成")
ws.getRange("B1:B1").setFormula("100*3")          # [5] 数式（先頭の "=" は付けない）
print("[6] getValue(True) -> B1 =", ws.getRange("B1:B1").getValue(True)["B1"])

wb.save()                                         # [7] 保存
wb.close()                                        # [8] 閉じる
print("[8] close() 完了")

wb = app.openWorkBook("flow-py.xlsx")             # [2'] 既存のブックを開く
ws = wb.openWorkSheet("Sheet1")
print("[2'] openWorkBook -> A1 =", ws.getRange("A1:A1").getValue(True)["A1"])
wb.close()
print("[8] close() 完了")