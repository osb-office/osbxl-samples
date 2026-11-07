# tables-comments.py — node 版サンプルと同一検証 (pyosbxl)
import json
import pyosbxl


def jdump(obj):  # JSON.stringify 相当の出力（区切り空白なし・非 ASCII そのまま）
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"))


app = pyosbxl.App()
wb = app.createWorkBook("tables-comments-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# ---------- データ投入（ヘッダ1行 + データ3行）----------
rows = [
    ["商品名", "単価", "数量"],
    ["リンゴ", "120", "3"],
    ["ミカン", "80", "5"],
    ["ブドウ", "300", "2"],
]
ws.setRangeData("A1", rows)

# ---------- 1. テーブルを作成 ----------
lo = ws.getListObjects()
table = lo.addList("Sales", "A1:C4", True, True)  # useFirstRowAsHeader, insertTotals
print("[1] テーブル作成")
print(f"    名前      = {table.getName()}")
print(f"    列名      = {jdump(table.getColumnName())}")
print(f"    スタイル  = {table.getTableStyleName()}")
print(f"    全範囲    = {table.getAllRange().getAddress()}")
print(f"    ヘッダ    = {table.getHeaderRowRange().getAddress()}")
print(f"    データ部  = {table.getDataBodyRange().getAddress()}")
print(f"    集計行    = {table.getTotalRowRange().getAddress()}")

# ---------- 2. 集計行 ----------
table.setTotalRowLabel(0, "合計")
table.setTotalRowFunction(1, pyosbxl.enums.XlTotalsCalculation.TotalsCalculationSum)
table.setTotalRowFunction(2, pyosbxl.enums.XlTotalsCalculation.TotalsCalculationSum)
print("[2] 集計行")
print(f"    ラベル(1列目 (0-based: 0)) = {jdump(table.getTotalRowLabel(1))}")
print(f"    数式          = {jdump(ws.getRange('B5:C5').getFormula())}")
print(f"    値            = {jdump(ws.getRange('B5:C5').getValue(True))}")

# ---------- 3. テーブルスタイルを変更 ----------
table.setBuiltinStyleName(pyosbxl.enums.XlDefaultTableStyle.TableStyleMedium9, True)
print(f"[3] スタイル変更後 = {table.getTableStyleName()}")

# ---------- 4. 構造化参照（テーブル列を参照する数式）----------
ws.getRange("E1:E1").setFormula("SUM(Sales[単価])")  # 先頭に = は付けない
print("[4] 構造化参照")
print(f"    E1 数式 = {jdump(ws.getRange('E1:E1').getFormula())}")
print(f"    E1 値   = {jdump(ws.getRange('E1:E1').getValue(True))}")

# ---------- 5. コメント（新式）とメモ（旧式）----------
cm = ws.getComments()


def mkComment(author, content):
    c = pyosbxl.dto.CommentObject()
    c.setAuthor(author)
    c.setContent(content)
    return c


# 新式コメント（スレッド）: 本体 + 返信の 2 件 → B2（save 後もファイルに残す）
cm.setCommentThread("B2", [mkComment("alice", "この数値を確認"), mkComment("bob", "確認しました")])
# 単一の新式コメント → D2（削除の実演用）
cm.setComment("D2", mkComment("dave", "あとで消すコメント"))
# 旧式メモ（吹き出し）→ A2
cm.setMemo("A2", "Carol", "最優先商品です", pyosbxl.dto.FontObject())

thread = cm.getComment("B2")
print("[5] コメント")
print(f"    B2 スレッド件数 = {len(thread)}")
for i, c in enumerate(thread):
    print(f"      [{i}] {c.getAuthor()}: {c.getContent()}{' (返信)' if c.getParentId() else ''}")
print(f"    D2 コメント件数 = {len(cm.getComment('D2'))}")
print(f"    A2 メモ = {jdump(cm.getMemoText('A2'))} / 作者 {jdump(cm.getMemoAuthor('A2'))}")

# 削除: 新式は removeComment / 旧式は removeMemo
cm.removeComment("D2")
cm.removeMemo("A2")


# 削除後の読み取りは node 版と同様に "comment not found" / "memo not found" 例外になる
def tryRead(label, fn):
    try:
        fn()
        print(f"    削除後 {label}: 例外なし")
    except Exception as e:
        print(f"    削除後 {label}: 例外 {e}")


tryRead("D2コメント", lambda: cm.getComment("D2"))
tryRead("A2メモ", lambda: cm.getMemoText("A2"))
print(f"    B2 は残すので save 後も {len(cm.getComment('B2'))}件")

total = ws.getRange("B5:B5").getValue(True)["B5"]
e1 = ws.getRange("E1:E1").getValue(True)["E1"]
print(f"[確認] テーブル={table.getName()} 集計(単価)={total} 構造化参照(E1)={e1}")

wb.save()
wb.close()