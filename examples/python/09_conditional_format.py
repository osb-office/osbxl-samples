# conditional-format.py — node 版サンプルと同一検証 (pyosbxl)
import pyosbxl

e = pyosbxl.enums


def color(hex):  # ColorObject は "#" なしの 6 桁 hex で作る
    c = pyosbxl.dto.ColorObject()
    c.setHexColor(hex)
    return c


def num(v):  # JS の数値と同じ表記にする（60.0 → "60"）
    return str(int(v)) if float(v) == int(v) else str(v)


def bol(v):  # JS の bool と同じ表記にする（True → "true"）
    return str(bool(v)).lower()


app = pyosbxl.App()
wb = app.createWorkBook("condfmt-py.xlsx")  # 出力文言は node 版に合わせ condfmt.xlsx と表示
ws = wb.openWorkSheet("Sheet1")

# ---------- データ表 ----------
# B列: セル値ルール / C列: カラースケール / D列: データバー / E列: 上位3件・平均比較
for i, h in enumerate(["月", "セル値ルール", "カラースケール", "データバー", "上位/平均"]):
    col = "ABCDE"[i]
    ws.getRange(f"{col}1:{col}1").setValue(h)
sales = [45, 78, 23, 91, 56, 34, 87, 12, 65, 49]
for i, v in enumerate(sales):
    ws.getRange(f"A{i + 2}:A{i + 2}").setNumberValue(i + 1)
    for col in ["B", "C", "D", "E"]:
        ws.getRange(f"{col}{i + 2}:{col}{i + 2}").setNumberValue(v)

fc = ws.getFormatConditions()

# ---------- (1) セル値ルール: 60 より大きい → 濃い赤の太字 + 薄赤の塗りつぶし ----------
cellVal = pyosbxl.dto.CellValueObject()
cellVal.setValueCondition(60, e.XlFormatConditionOperator.FormatConditionGreater)
cellVal.setA1C1("B2:B11")
redFont = pyosbxl.dto.FontObject()
redFont.setBold(True)
redFont.setColorObject(color("9C0006"))
cellVal.setFont(redFont)
pinkFill = pyosbxl.dto.PatternFillObject()
pinkFill.setPattern(e.XlPatternType.PatternSolid)
pinkFill.setCellColor(color("FFC7CE"))
cellVal.setPatternFill(pinkFill)
fc.setCellValueCondition(1, cellVal)

# ---------- (2) カラースケール: C列に 3 色のグラデーション ----------
colorScale = pyosbxl.dto.ColorScaleObject()
colorScale.setA1C1("C2:C11")
colorScale.setMinimumCondition(color("F8696B"), e.XlConditionValueType.ConditionValueTypeMinimum)
colorScale.setMediumCondition(color("FFEB84"), e.XlConditionValueType.ConditionValueTypePercentile, 50)
colorScale.setMaximumCondition(color("63BE7B"), e.XlConditionValueType.ConditionValueTypeMaximum)
fc.setColorScaleFormatCondition(2, colorScale)

# ---------- (3) データバー: D列 ----------
dataBar = pyosbxl.dto.DataBarObject()
dataBar.setA1C1("D2:D11")
dataBar.setDataBarColor(color("638EC6"))
dataBar.setMinimumCondition(e.XlConditionValueType.ConditionValueTypeMinimum)
dataBar.setMaximumCondition(e.XlConditionValueType.ConditionValueTypeMaximum)
fc.setDataBarFormatCondition(3, dataBar)

# ---------- (4) 上位 3 件 + 平均より上: E列 ----------
top3 = pyosbxl.dto.Top10Object()
top3.setRank(3, True)                        # True = 上位（False なら下位）
top3.setA1C1("E2:E11")
blueFont = pyosbxl.dto.FontObject()
blueFont.setBold(True)
blueFont.setColorObject(color("0070C0"))
top3.setFont(blueFont)
fc.setTop10FormatCondition(4, top3)

aboveAvg = pyosbxl.dto.AboveAverageObject()
aboveAvg.setAboveBelow(e.XlAboveBelow.XlAboveAverage)
aboveAvg.setA1C1("E2:E11")
greenFont = pyosbxl.dto.FontObject()
greenFont.setBold(True)
greenFont.setColorObject(color("006100"))
aboveAvg.setFont(greenFont)
fc.setAboveAverageCondition(5, aboveAvg)

fc.apply()                                   # 全条件の設定後に apply() が必須
print("[1] 5 件のルールを登録し apply() しました")

# ---------- ルール一覧の読み戻し ----------
types = fc.getConditionTypes()
print(f"[2] ルール数 = {len(types)}")
for prio in sorted(types):
    for rng, t in types[prio].items():
        print(f"    priority {prio}: {rng} → {t.name} (type={int(t)})")

# ---------- 各ルールの内容の読み戻し ----------
print("[3] ルール内容の読み戻し")
g1 = fc.getCellValueCondition(1)
print(f"    セル値  : {g1.getA1C1()} 値={num(g1.getConditionValue())}"
      f" 演算子={int(g1.getConditionOperator())}"
      f" (Greater={int(e.XlFormatConditionOperator.FormatConditionGreater)})"
      f" 文字色={g1.getFont().getColorObject().getHexColor()} 太字={bol(g1.getFont().isBold())}"
      f" 背景={g1.getPatternFill().getCellColor().getHexColor()}")
g2 = fc.getColorScaleFormatCondition(2)
print(f"    カラースケール: min={g2.getMinimumColor().getHexColor()} (type={int(g2.getMinimumValueType())})"
      f" mid={g2.getMediumColor().getHexColor()} (type={int(g2.getMediumValueType())}, 閾値={num(g2.getMediumValue())})"
      f" max={g2.getMaximumColor().getHexColor()} (type={int(g2.getMaximumValueType())})")
g3 = fc.getDataBarFormatCondition(3)
print(f"    データバー: 色={g3.getDataBarColor().getHexColor()}"
      f" minType={int(g3.getMinimumCondition())} maxType={int(g3.getMaximumCondition())}")
g4 = fc.getTop10FormatCondition(4)
print(f"    上位 N 件: rank={g4.getRank()} isTop={bol(g4.getTopBottom())}")
g5 = fc.getAboveAverageCondition(5)
print(f"    平均比較: aboveBelow={int(g5.getAboveBelow())}"
      f" (XlAboveAverage={int(e.XlAboveBelow.XlAboveAverage)})")

# ---------- 条件を満たして書式が付いたセル（apply() 後に取得可能）----------
byRow = lambda a: int(a[1:])
print("[4] 書式が適用されたセル")
appliedB = sorted(fc.getAppliedFontObject("B2:B11").keys(), key=byRow)
print(f"    B列 (>60)     : {' '.join(appliedB)}")
appliedE = sorted(fc.getAppliedFontObject("E2:E11").keys(), key=byRow)
print(f"    E列 (上位/平均): {' '.join(appliedE)}")
appliedBar = fc.getAppliedDataBar("D2:D11")
print(f"    D列 データバー : {sum(1 for v in appliedBar.values() if v)} セル")

# ---------- 削除 ----------
fc.removeCondition(4)                        # 上位 3 件のルールだけ削除
print(f"[5] removeCondition(4) 後のルール数 = {len(fc.getConditionTypes())}")
fc.removeAllCondition()                      # 全ルールを削除
print(f"[6] removeAllCondition() 後のルール数 = {len(fc.getConditionTypes())}")

# セル値ルールを登録し直して保存（保存ファイルを開いたとき効果が見えるように）
fc.setCellValueCondition(1, cellVal)
fc.apply()

wb.save()
wb.close()
print("[done] condfmt.xlsx を保存しました（セル値ルールを再登録済み）")