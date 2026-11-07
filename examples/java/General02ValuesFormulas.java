// values-formulas.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

public class General02ValuesFormulas {
    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("sales-java.xlsx");   // 新規ブック作成
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");             // 既定のシートを開く

        // --- 見出し (1行目) ---
        // app.convertFromRowColNumber(行, 列) で (1,1) -> "A1" のセルアドレスが得られる
        String[] headers = { "商品", "単価", "個数", "金額" };
        for (int i = 0; i < headers.length; i++) {
            String cell = app.convertFromRowColNumber(1, i + 1);
            ws.getRange(cell + ":" + cell).setValue(headers[i]);
        }

        // --- 明細 (2〜4行目) ---
        String[] names = { "りんご", "みかん", "ぶどう" };
        double[] prices = { 100, 50, 150 };
        double[] qtys = { 3, 6, 2 };
        for (int i = 0; i < names.length; i++) {
            int row = i + 2;
            ws.getRange("A" + row + ":A" + row).setValue(names[i]);
            ws.getRange("B" + row + ":B" + row).setNumberValue(prices[i]);
            ws.getRange("C" + row + ":C" + row).setNumberValue(qtys[i]);
        }

        // --- 金額列 (D2:D4) ---
        // 行ごとに異なる相対参照を入れたいので setRangeFormulas を使う。
        // 数式文字列に先頭の "=" は付けない。
        List<List<String>> formulas = new ArrayList<>();
        for (int i = 0; i < names.length; i++) {
            formulas.add(Collections.singletonList("B" + (i + 2) + "*C" + (i + 2)));
        }
        ws.setRangeFormulas("D2", formulas);

        // --- 合計 (D5) ---
        ws.getRange("C5:C5").setValue("合計");
        ws.getRange("D5:D5").setFormula("SUM(D2:D4)");

        // --- 読み戻し ---
        // getValue(true) は「生値」= 数式の計算結果を返す
        Map<String, String> values = ws.getRange("A1:D5").getValue(true);
        System.out.println("金額 D2:D4 = "
                + String.join(", ", values.get("D2"), values.get("D3"), values.get("D4")));
        System.out.println("合計: " + values.get("D5"));

        // --- 保存して閉じる ---
        wb.save();
        wb.close();
    }
}