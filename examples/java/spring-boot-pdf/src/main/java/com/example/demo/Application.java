package com.example.demo;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;

import com.osboffice.osbxl.*;
import com.osboffice.osbxl.chart.*;
import com.osboffice.osbxl.enums.*;
import com.osboffice.osbxl.dto.*;

import java.util.Collections;

@SpringBootApplication
@RestController
public class Application {

    public static void main(String[] args) {
        SpringApplication.run(Application.class, args);
    }

    @PostMapping("/api/export-pdf")
    public ResponseEntity<byte[]> exportPdf() {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("sales_report.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet(0); // 0 始まり (zeroBase)

        // 1. タイトル
        ws.getRange("A1").setValue("2026年度 四半期別 売上実績レポート");

        // 2. テーブルヘッダー (3行目: 行インデックス 2)
        String[] headers = {"四半期", "商品A (万円)", "商品B (万円)", "合計 (万円)"};
        for (int c = 0; c < headers.length; c++) {
            ws.getCells(2, c).setValue(headers[c]);
        }

        // 3. 明細データと合計計算式 (4〜7行目: 行インデックス 3〜6)
        String[] quarters = {"第1四半期", "第2四半期", "第3四半期", "第4四半期"};
        double[][] sales = {
            {120, 95},
            {145, 110},
            {160, 125},
            {190, 140}
        };

        for (int r = 0; r < quarters.length; r++) {
            int rowIdx = 3 + r;
            ws.getCells(rowIdx, 0).setValue(quarters[r]);
            ws.getCells(rowIdx, 1).setNumberValue(sales[r][0]);
            ws.getCells(rowIdx, 2).setNumberValue(sales[r][1]);
            int excelRow = rowIdx + 1; // 1-based 表示行
            ws.getCells(rowIdx, 3).setFormula("SUM(B" + excelRow + ":C" + excelRow + ")");
        }

        // 4. 構造化テーブル作成
        ListObjectsWrapper lo = ws.getListObjects();
        TableWrapper table = lo.addList("SalesTable", "A3:D7", true, true);
        table.setTotalRowLabel(0, "通期合計");
        table.setTotalRowFunction(1, XlTotalsCalculation.TotalsCalculationSum);
        table.setTotalRowFunction(2, XlTotalsCalculation.TotalsCalculationSum);
        table.setTotalRowFunction(3, XlTotalsCalculation.TotalsCalculationSum);

        // 5. 集合縦棒グラフの配置 (left: 380pt, top: 40pt)
        ChartObjectWrapper co = ws.addChartObject("SalesChart", 380, 40, 420, 240);
        com.osboffice.osbxl.chart.ChartWrapper chart = co.getChart();
        chart.chartWizard(
            "Sheet1!A3:C7",
            XlChartType.ChartTypeColumnClustered,
            XlRowCol.Columns,
            1, 1, 1,
            XlChartColorPalette.XlPaletteColorful1,
            true,
            "四半期売上推移 (商品別)"
        );

        // 6. PageSetupObject による印刷・PDF ページレイアウト設定
        PageSetupObjectWrapper ps = ws.getPageSetupObject();
        ps.setPaperSize(XlPaperSize.PaperA4);
        ps.setOrientation(XlPageOrientation.Landscape);
        ps.setCenterHorizontally(true);
        ps.setCenterVertically(true);
        ps.setFitToPages(1, 1);

        HeaderObjectWrapper header = new HeaderObjectWrapper();
        header.setLeftHeader("株式会社サンプル");
        header.setRightHeader("出力日: 2026/11/07");
        ps.setHeader(header);

        FooterObjectWrapper footer = new FooterObjectWrapper();
        footer.setCenterFooter("ページ 1 / 1");
        ps.setFooter(footer);
        ws.setPageSetupObject(ps);

        // 7. PdfExportObject によるセキュリティ・出力オプション設定
        PdfExportObjectWrapper pdfOpt = new PdfExportObjectWrapper();
        pdfOpt.setOpenPassword("customer2026");       // 閲覧パスワード
        pdfOpt.setPermissionsPassword("adminSecret"); // 権限パスワード
        pdfOpt.setAllowPrint(true);
        pdfOpt.setAllowCopy(false);
        pdfOpt.setAllowModify(false);
        pdfOpt.setColorMode(XlPdfColorMode.Color);
        pdfOpt.setFitMode(XlPdfFitMode.FitAllColumnsOnOnePage);

        // 8. メモリから直接 byte[] を取得 (ファイル一時保存なし)
        byte[] pdfBytes = wb.saveAsPdfToBuffer(false, Collections.singletonList("Sheet1"), true, pdfOpt);
        wb.close();

        // 9. HTTP レスポンスとしてバイナリストリーミング返却
        return ResponseEntity.ok()
                .contentType(MediaType.APPLICATION_PDF)
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"sales_report_spring.pdf\"")
                .header(HttpHeaders.CONTENT_LENGTH, String.valueOf(pdfBytes.length))
                .body(pdfBytes);
    }
}
