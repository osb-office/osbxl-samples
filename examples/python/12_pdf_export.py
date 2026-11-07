# pdf_export.py — PDF エクスポートとページ設定 (pyosbxl)
import pyosbxl
from pyosbxl import enums

app = pyosbxl.App()
wb = app.createWorkBook("report.xlsx")
ws = wb.openWorkSheet(0)  # 0 始まり (zeroBase)

# 1. レポート表の作成
ws.getRange("A1").setValue("2026年度 部門別売上実績レポート")
headers = ["部門名", "第1四半期", "第2四半期", "第3四半期", "第4四半期", "合計"]
for c, h in enumerate(headers):
    ws.getCells(2, c).setValue(h)

rows = [
    ["営業1課", 1200000, 1450000, 1380000, 1820000],
    ["営業2課", 980000,  1120000, 1050000, 1280000],
    ["開発営業", 2150000, 2300000, 2420000, 2890000]
]
for r, row in enumerate(rows):
    ws.getCells(3 + r, 0).setValue(row[0])
    for c in range(1, 5):
        ws.getCells(3 + r, c).setNumberValue(row[c])
    ws.getCells(3 + r, 5).setFormula(f"SUM(B{4 + r}:E{4 + r})")

# 2. ページ設定（A4・横向き・余白調整・水平中央配置）
ps = ws.getPageSetupObject()
ps.setPaperSize(enums.XlPaperSize.PaperA4)
ps.setOrientation(enums.XlPageOrientation.Landscape)
ps.setLeftMargin(0.5)
ps.setRightMargin(0.5)
ps.setTopMargin(0.6)
ps.setBottomMargin(0.6)
ps.setCenterHorizontally(True)
ws.setPageSetupObject(ps)

# 3. ブックの保存
wb.save()

# 4. PDF 出力（ファイル出力）
pdf_result = wb.saveAsPdf("report.pdf", False)  # includeLayout = False
print(f"PDF 出力完了: 総ページ数 = {pdf_result.getPageCount()}")

# 5. メモリバッファへの出力
pdf_buffer = wb.saveAsPdfToBuffer(False, ["Sheet1"])
print(f"PDF バッファ生成完了: バイト長 = {len(pdf_buffer)}")

wb.close()