# sheets.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("sheets-py.xlsx")

def b(x):
    # python の True/False を node の true/false 表記に合わせる
    return "true" if x else "false"

# インデックスは 0 始まり (zeroBase)。0..getSheetCount()-1 を順に開いて名前を並べる
def show(label):
    names = [wb.openWorkSheet(i).getName() for i in range(wb.getSheetCount())]
    print(f"{label} count={wb.getSheetCount()} [{', '.join(names)}]")

show("[1] 新規作成直後        :")

wb.addWorkSheet("集計")                       # position 省略 → 末尾に追加
show("[2] addWorkSheet(集計)  :")

wb.addWorkSheet("受注", 0)                    # position=0 → 先頭に挿入 (0-based)
show("[3] addWorkSheet(受注,0):")

wb.moveSheet("受注", 2)                       # 3 番目へ移動 (0-based: index 2)
show("[4] moveSheet(受注,2)   :")

ws = wb.openWorkSheet("受注")
print("[5] getName() =", ws.getName(), "/ getSheetNumber(受注) =", wb.getSheetNumber("受注"))
ws.getRange("A1:A1").setValue("受注データ")

wb.openWorkSheet("Sheet1").setName("明細")    # リネーム
show("[6] setName(明細)       :")

# --- アクティブタブ（ブックを開いたときに手前に表示されるシート）---
print("[7] isActive: 集計 =", b(wb.openWorkSheet("集計").isActive()),
      "/ 明細 =", b(wb.openWorkSheet("明細").isActive()))
wb.openWorkSheet("集計").activate(True)       # 他のシートのアクティブは自動的に解除されます
print("    activate後: 集計 =", b(wb.openWorkSheet("集計").isActive()),
      "/ 明細 =", b(wb.openWorkSheet("明細").isActive()))

# --- 削除 ---
wb.addWorkSheet("一時")
wb.deleteWorkSheet("一時")
show("[8] deleteWorkSheet(一時):")

wb.save()
wb.close()

# --- 開き直して確認 ---
wb2 = app.openWorkBook("sheets-py.xlsx")
names = [wb2.openWorkSheet(i).getName() for i in range(wb2.getSheetCount())]
print("[終] シート構成:", ", ".join(names))
print("     受注!A1 =", wb2.openWorkSheet("受注").getRange("A1:A1").getValue()["A1"])
wb2.close()