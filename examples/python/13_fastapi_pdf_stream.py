"""13_fastapi_pdf_stream.py — FastAPI によるインメモリ PDF 直接ストリーミング (pyosbxl)

ディスクへの一時ファイル保存を一切行わず、wb.saveAsPdfToBuffer() で生成した
PDF バイナリ (bytes) を FastAPI の Response で直接クライアントへ返却します。
"""
from fastapi import FastAPI, Response
import uvicorn
import pyosbxl
from pyosbxl import dto, enums

app = FastAPI(title="osbxl FastAPI PDF Streaming Service")

@app.post("/api/export-pdf")
def export_pdf():
    # 1. ワークブックとシートの準備 (zeroBase: 0-based)
    xl_app = pyosbxl.App()
    wb = xl_app.createWorkBook("sales_report.xlsx")
    ws = wb.openWorkSheet(0)

    # 2. 表タイトルとヘッダー
    ws.getRange("A1").setValue("2026年度 四半期別 売上実績レポート")
    headers = ["四半期", "商品A (万円)", "商品B (万円)", "合計 (万円)"]
    for c, h in enumerate(headers):
        ws.getCells(2, c).setValue(h)

    # 3. 明細データと合計計算式 (4〜7行目: 行インデックス 3〜6)
    sales = [
        ["第1四半期", 120,  95],
        ["第2四半期", 145, 110],
        ["第3四半期", 160, 125],
        ["第4四半期", 190, 140],
    ]
    for idx, row in enumerate(sales):
        r = 3 + idx
        ws.getCells(r, 0).setValue(row[0])
        ws.getCells(r, 1).setNumberValue(row[1])
        ws.getCells(r, 2).setNumberValue(row[2])
        excel_row = r + 1  # 1-based 表示行
        ws.getCells(r, 3).setFormula(f"SUM(B{excel_row}:C{excel_row})")

    # 4. 構造化テーブル (ListObjects) と集計行 (SUM) の作成
    lo = ws.getListObjects()
    table = lo.addList("SalesTable", "A3:D7", True, True)
    table.setTotalRowLabel(0, "通期合計")
    table.setTotalRowFunction(1, enums.XlTotalsCalculation.TotalsCalculationSum)
    table.setTotalRowFunction(2, enums.XlTotalsCalculation.TotalsCalculationSum)
    table.setTotalRowFunction(3, enums.XlTotalsCalculation.TotalsCalculationSum)

    # 5. 集合縦棒グラフの配置 (left: 380pt, top: 40pt)
    co = ws.addChartObject("SalesChart", 380, 40, 420, 240)
    chart = co.getChart()
    chart.chartWizard(
        "Sheet1!A3:C7",
        enums.XlChartType.ChartTypeColumnClustered,
        enums.XlRowCol.Columns,
        1, 1, 1,
        enums.XlChartColorPalette.XlPaletteColorful1,
        True,
        "四半期売上推移 (商品別)"
    )

    # 6. PageSetupObject による印刷・PDF ページレイアウト設定
    ps = ws.getPageSetupObject()
    ps.setPaperSize(enums.XlPaperSize.PaperA4)
    ps.setOrientation(enums.XlPageOrientation.Landscape)
    ps.setCenterHorizontally(True)
    ps.setCenterVertically(True)
    ps.setFitToPages(1, 1)

    header = dto.HeaderObject()
    header.setLeftHeader("株式会社サンプル")
    header.setRightHeader("出力日: 2026/11/07")
    ps.setHeader(header)

    footer = dto.FooterObject()
    footer.setCenterFooter("ページ 1 / 1")
    ps.setFooter(footer)
    ws.setPageSetupObject(ps)

    # 7. PdfExportObject によるセキュリティ・出力オプション設定
    pdf_opt = dto.PdfExportObject()
    pdf_opt.setOpenPassword("customer2026")       # 開封パスワード (AES-256)
    pdf_opt.setPermissionsPassword("adminSecret") # 権限マスターパスワード
    pdf_opt.setAllowPrint(True)                   # 印刷許可
    pdf_opt.setAllowCopy(False)                  # コピー禁止
    pdf_opt.setAllowModify(False)                # 文書改変禁止
    pdf_opt.setColorMode(enums.XlPdfColorMode.Color)
    pdf_opt.setFitMode(enums.XlPdfFitMode.FitAllColumnsOnOnePage)

    # 8. ディスク保存なしでメモリから直接 PDF バイト列を取得
    pdf_bytes = wb.saveAsPdfToBuffer(False, ["Sheet1"], True, pdf_opt)
    wb.close()

    # 9. HTTP レスポンスとして直接ストリーミング返却
    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={
            "Content-Disposition": 'attachment; filename="sales_report_fastapi.pdf"',
            "Content-Length": str(len(pdf_bytes))
        }
    )

if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8080)
