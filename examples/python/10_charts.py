# charts.py — チャート（グラフ）の基本: node 版サンプルと同一検証 (pyosbxl)
import json
import pyosbxl

enums = pyosbxl.enums

app = pyosbxl.App()
wb = app.createWorkBook("charts-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# 補助: JSON 相当の文字列（区切りスペース無し・非 ASCII はそのまま）
def jdump(v):
    return json.dumps(v, ensure_ascii=False, separators=(",", ":"))

# 補助: enum のメンバー名（pyosbxl の enum は __name__ で名前が取れる）
def enum_name(v):
    return v.__name__

# 補助: シート上のチャート個数（getChartObject は範囲外で例外 → 数えて判定）
def count_charts():
    n = 0
    while True:
        try:
            ws.getChartObject(n)
            n += 1
        except Exception:
            return n

# 補助: チャートの系列数（getSeriesCollection は範囲外で例外）
def count_series(chart):
    n = 0
    while True:
        try:
            chart.getSeriesCollection(n + 1)
            n += 1
        except Exception:
            return n

# ---------- 1. データ表（2商品の月次売上）を作る ----------
for c, h in enumerate(["月", "商品A", "商品B"]):
    ws.getCells(1, c + 1).setValue(h)
sales = [
    ["1月", 120, 90], ["2月", 135, 95], ["3月", 150, 110],
    ["4月", 142, 125], ["5月", 160, 118], ["6月", 175, 130],
]
for i, row in enumerate(sales):
    ws.getCells(i + 2, 1).setValue(row[0])
    ws.getCells(i + 2, 2).setNumberValue(row[1])
    ws.getCells(i + 2, 3).setNumberValue(row[2])
print("[1] データ表 A1:C7 を作成 / 商品A合計 =", sum(r[1] for r in sales))

# ---------- 2. 集合縦棒グラフを追加（chartWizard）----------
co1 = ws.addChartObject("SalesColumn", 320, 10, 360, 220)
chart1 = co1.getChart()
chart1.chartWizard(
    "Sheet1!A1:C7",
    enums.XlChartType.ChartTypeColumnClustered,
    enums.XlRowCol.Columns,   # 系列を列方向に取る（商品A/商品B が 1 系列ずつ）
    1,                        # seriesLabelLines: 先頭 1 行が系列名ヘッダ
    1,                        # categoryLabelLines: 先頭 1 列がカテゴリ名ヘッダ
    1,                        # chartStyleNo（0 始まり (zeroBase)）
    enums.XlChartColorPalette.XlPaletteColorful1,
    True,                     # hasLegend: 凡例を表示
    "月次売上",               # title: 設定するとタイトルが表示される
)
print("[2] 型 =", enum_name(chart1.getChartType()),
      "/ hasTitle =", str(chart1.hasTitle()).lower(),
      "/ タイトル =", jdump(chart1.getChartTitle().getText()))
print("    データ源 =", chart1.getDataSource(),
      "/ plotBy =", int(chart1.getPlotBy()),
      "/ hasLegend =", str(chart1.hasLegend()).lower())

# ---------- 3. 系列・軸の読み戻し ----------
n1 = count_series(chart1)
print("[3] 系列数 =", n1)
for i in range(1, n1 + 1):
    s = chart1.getSeriesCollection(i)
    name = list(ws.getCells(1, i + 1).getValue().values())[0]  # 系列 i → 列 i+1 のヘッダ
    print(f"    系列{i}「{name}」値 =", jdump(s.getValues()),
          "/ ポイント数 =", s.getPointSize())
print("    カテゴリ名 =",
      jdump(chart1.getAxes(enums.XlAxisType.Category).getCategoryNames()))

# ---------- 4. 2つ目のグラフ（マーカー付き折れ線）+ 軸・凡例の設定 ----------
co2 = ws.addChartObject("SalesLine", 320, 250, 360, 220)
chart2 = co2.getChart()
chart2.chartWizard(
    "Sheet1!A1:C7",
    enums.XlChartType.ChartTypeLineMarkers,
    enums.XlRowCol.Columns, 1, 1,
    1, enums.XlChartColorPalette.XlPaletteColorful1, True)  # hasLegend=True にしないと getLegend() が例外
chart2.setTitle(True)
chart2.getChartTitle().setText("月次売上推移")
val_axis = chart2.getAxes(enums.XlAxisType.Value)
val_axis.setMaximumScale(200)
val_axis.setMinimumScale(0)
val_axis.setMajorUnit(50)
val_axis.setTitle(True)
val_axis.getTitle().setText("売上(万円)")
chart2.getLegend().setPosition(enums.XlLegendPosition.LegendPositionRight)

print("[4] 型 =", enum_name(chart2.getChartType()),
      "/ タイトル =", jdump(chart2.getChartTitle().getText()))
print("    値軸 主単位 =", int(val_axis.getMajorUnit()),
      "/ 最大自動 =", str(val_axis.getMaximumScaleIsAuto()).lower(),
      "/ 最小自動 =", str(val_axis.getMinimumScaleIsAuto()).lower())
print("    値軸タイトル =", jdump(val_axis.getTitle().getText()),
      "/ 凡例位置 =", int(chart2.getLegend().getPosition()),
      f"(Right={int(enums.XlLegendPosition.LegendPositionRight)})")

# ---------- 5. チャート個数・位置・サイズの読み戻し ----------
print("[5] チャート個数 =", count_charts(),
      "/ SalesColumn(left,top,width,height) =",
      f"{int(co1.getLeft())},{int(co1.getTop())},{int(co1.getWidth())},{int(co1.getHeight())}")

# ---------- 6. チャートの削除 ----------
ws.deleteChartObject("SalesLine")   # 名前指定。deleteChartObject(2) のように索引(1始まり)でも可
print("[完了] 削除後のチャート個数 =", count_charts(),
      "/ 残 =", ws.getChartObject(0).getName())

wb.save()
wb.close()