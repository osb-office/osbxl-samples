// PdfExportExample.java — PDF エクスポートとページ設定
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.enums.*;
import com.osboffice.osbxl.dto.*;
import java.util.Collections;

public class General12PdfExport {
    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("report-java.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet(0); // 0 始まり (zeroBase)

        // 1. レポート表の作成
        ws.getRange("A1").setValue("2026年度 部門別売上実績レポート");
        String[] headers = {"部門名", "第1四半期", "第2四半期", "第3四半期", "第4四半期", "合計"};
        for (int c = 0; c < headers.length; c++) {
            ws.getCells(2, c).setValue(headers[c]);
        }

        String[] depts = {"営業1課", "営業2課", "開発営業"};
        double[][] sales = {
            {1200000, 1450000, 1380000, 1820000},
            {980000,  1120000, 1050000, 1280000},
            {2150000, 2300000, 2420000, 2890000}
        };

        for (int r = 0; r < depts.length; r++) {
            ws.getCells(3 + r, 0).setValue(depts[r]);
            for (int c = 0; c < 4; c++) {
                ws.getCells(3 + r, c + 1).setNumberValue(sales[r][c]);
            }
            ws.getCells(3 + r, 5).setFormula("SUM(B" + (4 + r) + ":E" + (4 + r) + ")");
        }

        // 2. ページ設定
        PageSetupObjectWrapper ps = ws.getPageSetupObject();
        ps.setPaperSize(XlPaperSize.PaperA4);
        ps.setOrientation(XlPageOrientation.Landscape);
        ps.setLeftMargin(0.5);
        ps.setRightMargin(0.5);
        ps.setTopMargin(0.6);
        ps.setBottomMargin(0.6);
        ps.setCenterHorizontally(true);
        ws.setPageSetupObject(ps);

        // 3. ブックの保存
        wb.save();

        // 4. PDF 出力
        PdfExportResultWrapper result = wb.saveAsPdf("report-java.pdf", false);
        System.out.println("PDF 出力完了: 総ページ数 = " + result.getPageCount());

        // 5. メモリバッファへの出力
        byte[] buffer = wb.saveAsPdfToBuffer(false, Collections.singletonList("Sheet1"));
        System.out.println("PDF バッファ生成完了: バイト長 = " + buffer.length);

        wb.close();
    }
}