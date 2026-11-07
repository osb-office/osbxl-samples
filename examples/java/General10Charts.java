// charts.java — チャート（グラフ）の基本: node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.chart.*;
import com.osboffice.osbxl.enums.*;
import java.util.List;
import java.util.Map;

public class General10Charts {
    // 補助: 文字列リスト → JSON 相当（["a","b"] 形式。node の JSON.stringify と同じ見た目にする）
    private static String jsonArr(List<String> list) {
        StringBuilder sb = new StringBuilder("[");
        for (int i = 0; i < list.size(); i++) {
            if (i > 0) sb.append(",");
            sb.append("\"").append(list.get(i)).append("\"");
        }
        return sb.append("]").toString();
    }

    // 補助: シート上のチャート個数（getChartObject は範囲外で例外 → 数えて判定）
    private static int countCharts(WorkSheetWrapper ws) {
        int n = 0;
        while (true) {
            try { ws.getChartObject(n); n++; } catch (Throwable t) { break; }
        }
        return n;
    }

    // 補助: チャートの系列数（getSeriesCollection は範囲外で例外）
    private static int countSeries(ChartWrapper chart) {
        int n = 0;
        while (true) {
            try { chart.getSeriesCollection(n + 1); n++; } catch (Throwable t) { break; }
        }
        return n;
    }

    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("charts-java.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");

        // ---------- 1. データ表（2商品の月次売上）を作る ----------
        String[] headers = {"月", "商品A", "商品B"};
        for (int c = 0; c < headers.length; c++) ws.getCells(1, c + 1).setValue(headers[c]);
        String[] months = {"1月", "2月", "3月", "4月", "5月", "6月"};
        int[] salesA = {120, 135, 150, 142, 160, 175};
        int[] salesB = {90, 95, 110, 125, 118, 130};
        int totalA = 0;
        for (int i = 0; i < months.length; i++) {
            ws.getCells(i + 2, 1).setValue(months[i]);
            ws.getCells(i + 2, 2).setNumberValue(salesA[i]);
            ws.getCells(i + 2, 3).setNumberValue(salesB[i]);
            totalA += salesA[i];
        }
        System.out.println("[1] データ表 A1:C7 を作成 / 商品A合計 = " + totalA);

        // ---------- 2. 集合縦棒グラフを追加（chartWizard）----------
        ChartObjectWrapper co1 = ws.addChartObject("SalesColumn", 320, 10, 360, 220);
        ChartWrapper chart1 = co1.getChart();
        chart1.chartWizard(
            "Sheet1!A1:C7",
            XlChartType.ChartTypeColumnClustered,
            XlRowCol.Columns,   // 系列を列方向に取る（商品A/商品B が 1 系列ずつ）
            1,                  // seriesLabelLines: 先頭 1 行が系列名ヘッダ
            1,                  // categoryLabelLines: 先頭 1 列がカテゴリ名ヘッダ
            1,                  // chartStyleNo（0 始まり (zeroBase)）
            XlChartColorPalette.XlPaletteColorful1,
            true,               // hasLegend: 凡例を表示
            "月次売上");         // title: 設定するとタイトルが表示される
        System.out.println("[2] 型 = " + chart1.getChartType().name()
            + " / hasTitle = " + chart1.hasTitle()
            + " / タイトル = \"" + chart1.getChartTitle().getText() + "\"");
        System.out.println("    データ源 = " + chart1.getDataSource()
            + " / plotBy = " + chart1.getPlotBy().getCode()   // enum は getCode() で数値にする
            + " / hasLegend = " + chart1.hasLegend());

        // ---------- 3. 系列・軸の読み戻し ----------
        int n1 = countSeries(chart1);
        System.out.println("[3] 系列数 = " + n1);
        for (int i = 1; i <= n1; i++) {
            SeriesWrapper s = chart1.getSeriesCollection(i);
            Map<String, String> h = ws.getCells(1, i + 1).getValue();
            String name = h.values().iterator().next();       // 系列 i → 列 i+1 のヘッダ
            System.out.println("    系列" + i + "「" + name + "」値 = " + jsonArr(s.getValues())
                + " / ポイント数 = " + s.getPointSize());
        }
        System.out.println("    カテゴリ名 = "
            + jsonArr(chart1.getAxes(XlAxisType.Category).getCategoryNames()));

        // ---------- 4. 2つ目のグラフ（マーカー付き折れ線）+ 軸・凡例の設定 ----------
        ChartObjectWrapper co2 = ws.addChartObject("SalesLine", 320, 250, 360, 220);
        ChartWrapper chart2 = co2.getChart();
        chart2.chartWizard("Sheet1!A1:C7",
            XlChartType.ChartTypeLineMarkers,
            XlRowCol.Columns, 1, 1,
            1, XlChartColorPalette.XlPaletteColorful1, true); // hasLegend=true にしないと getLegend() が例外
        chart2.setTitle(true);
        chart2.getChartTitle().setText("月次売上推移");
        AxisWrapper valAxis = chart2.getAxes(XlAxisType.Value);
        valAxis.setMaximumScale(200);
        valAxis.setMinimumScale(0);
        valAxis.setMajorUnit(50);
        valAxis.setTitle(true);
        valAxis.getTitle().setText("売上(万円)");
        chart2.getLegend().setPosition(XlLegendPosition.LegendPositionRight);

        System.out.println("[4] 型 = " + chart2.getChartType().name()
            + " / タイトル = \"" + chart2.getChartTitle().getText() + "\"");
        System.out.println("    値軸 主単位 = " + (int) valAxis.getMajorUnit()   // double → int で表示
            + " / 最大自動 = " + valAxis.getMaximumScaleIsAuto()
            + " / 最小自動 = " + valAxis.getMinimumScaleIsAuto());
        System.out.println("    値軸タイトル = \"" + valAxis.getTitle().getText() + "\""
            + " / 凡例位置 = " + chart2.getLegend().getPosition().getCode()
            + " (Right=" + XlLegendPosition.LegendPositionRight.getCode() + ")");

        // ---------- 5. チャート個数・位置・サイズの読み戻し ----------
        System.out.println("[5] チャート個数 = " + countCharts(ws)
            + " / SalesColumn(left,top,width,height) = "
            + (int) co1.getLeft() + "," + (int) co1.getTop() + ","
            + (int) co1.getWidth() + "," + (int) co1.getHeight());

        // ---------- 6. チャートの削除 ----------
        ws.deleteChartObject("SalesLine");   // 名前指定。deleteChartObject(2) のように索引(1始まり)でも可
        System.out.println("[完了] 削除後のチャート個数 = " + countCharts(ws)
            + " / 残 = " + ws.getChartObject(0).getName());

        wb.save();
        wb.close();
    }
}