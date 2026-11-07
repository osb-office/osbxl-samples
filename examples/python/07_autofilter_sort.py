# af-sort.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("af-sort-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# ---------- 1. データ準備（ヘッダ + 10 行）----------
DATA = [
    ["りんご", "東京", 120, 100],
    ["みかん", "大阪", 40, 80],
    ["バナナ", "東京", 200, 150],
    ["ぶどう", "福岡", 60, 300],
    ["なし", "大阪", 90, 120],
    ["もも", "福岡", 150, 400],
    ["キウイ", "東京", 30, 90],
    ["マンゴー", "大阪", 70, 500],
    ["いちご", "福岡", 110, 350],
    ["メロン", "東京", 25, 800],
]
for i, h in enumerate(["商品名", "地域", "数量", "単価"]):
    col = "ABCD"[i]
    ws.getRange(f"{col}1:{col}1").setValue(h)
for i, row in enumerate(DATA):
    r = i + 2
    ws.getRange(f"A{r}:A{r}").setValue(row[0])
    ws.getRange(f"B{r}:B{r}").setValue(row[1])
    ws.getRange(f"C{r}:C{r}").setNumberValue(row[2])
    ws.getRange(f"D{r}:D{r}").setNumberValue(row[3])
print("[1] データ準備: A1:D11（ヘッダ + 10 行）")

# 表示用ヘルパー: 行番号 → "商品名(値)" / 非表示でない行の一覧
def text(r, col):
    return ws.getRange(f"{col}{r}:{col}{r}").getValue()[f"{col}{r}"]

def label(r, col):
    return f"{text(r, 'A')}({text(r, col)})"

def visible(col):
    out = []
    for r in range(2, 12):
        if not ws.getRow(r).getHidden():
            out.append(label(r, col))
    return " ".join(out)

def hidden_list():
    out = [str(r) for r in range(2, 12) if ws.getRow(r).getHidden()]
    return ",".join(out) if out else "なし"

# ---------- 2. オートフィルター設定 ----------
af = ws.getAutoFilter("A1:D11", True)   # 先頭行はヘッダ
print(f"[2] getAutoFilter → 範囲 = {af.getAddress()} / 非表示行 = {hidden_list()}")

# ---------- 3. 条件フィルター: 数量 > 100 ----------
# colId は範囲内の列番号（A=1 … D=4）。criteria に演算子記号は含めない
af.setCustomFilter(pyosbxl.enums.XlAutoFilterOperator.GreaterThan, "100", 3)
print(f"[3] 数量>100 の可視行: {visible('C')}")
print(f'    非表示行 = {hidden_list()} / getValue("A2:A11") の件数 = {len(ws.getRange("A2:A11").getValue())}')
af.resetAllFilter()                     # 条件だけ解除（フィルター自体は残る）

# ---------- 4. Top-N / 値リストのフィルター ----------
af.setTop10ValueFilter(3, 3)            # 数量の上位 3 件
print(f"[4] 数量上位3件の可視行: {visible('C')}")
af.resetAllFilter()
af.setUniqueValuesNumberFilter([120, 40], 3)   # 数量が 120 または 40
print(f"[5] 数量∈{{120,40}} の可視行: {visible('C')}")
af.resetAllFilter()
af.setUniqueValuesStringFilter(["東京", "福岡"], 2)   # 地域が 東京 または 福岡（文字列リスト）
print(f"[6] 地域∈{{東京,福岡}} の可視行: {visible('B')}")
af.resetAllFilter()
af.setCustomFilter(pyosbxl.enums.XlAutoFilterOperator.Equal, "大阪", 2)   # 地域 = 大阪（文字列一致）
print(f'[7] 地域="大阪" の可視行: {visible("B")}')
af.resetAllFilter()

# ---------- 5. 並べ替え（単一キー）----------
# ヘッダを含まないデータ部だけを範囲にする（firstRowAsHeader 省略 = False）
ws.getSort("A2:D11").executeSortAscending(2)   # 数量（範囲内3列目 (0-based: 2)）で昇順
print(f"[8] 数量の昇順: {visible('C')}")

# オートフィルター経由の Sort はヘッダ行が自動的に保護される
sort = af.getSort()
sort.executeSortDescending(3)                  # 単価（4列目 (0-based: 3)）で降順
print(f"[9] 単価の降順: {visible('D')}")

# ---------- 6. 並び替え条件の読み戻し ----------
cond = sort.getSortConditions()[0]
# Python の bool は True/False 表示なので lower() で JS 風に揃える
print(f"[10] 条件の読み戻し: 対象列 = {cond.getTarget()} / 昇順 = {str(cond.getSortAscending()).lower()}")

# ---------- 7. オートフィルター自体の削除 ----------
af.removeFilter()
print(f"[11] removeFilter 後: 非表示行 = {hidden_list()}")

wb.save()
wb.close()