// pivot-sample.java — ピボットテーブル: node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import com.osboffice.osbxl.enums.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class General11Pivottables {
    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("pivot-java.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");

        // ---------- 1. ソースデータ ----------
        // 地域 / 商品 / 数量 の 10 行
        String[] headers = {"地域", "商品", "数量"};
        String[][] srcData = {
            {"東", "りんご", "10"}, {"東", "みかん", "20"}, {"西", "りんご", "30"},
            {"西", "みかん", "15"}, {"東", "りんご", "25"}, {"北", "みかん", "40"},
            {"北", "りんご",  "5"}, {"西", "りんご", "10"}, {"東", "みかん", "30"},
            {"北", "りんご", "20"},
        };
        for (int i = 0; i < headers.length; i++) {
            String col = String.valueOf((char) ('A' + i));
            ws.getRange(col + "1:" + col + "1").setValue(headers[i]);
        }
        int srcTotal = 0;
        for (int ri = 0; ri < srcData.length; ri++) {
            for (int ci = 0; ci < srcData[ri].length; ci++) {
                String col = String.valueOf((char) ('A' + ci));
                int r = ri + 2;
                String addr = col + r + ":" + col + r;
                if (ci == 2) {   // 数量列は数値として入力
                    int v = Integer.parseInt(srcData[ri][ci]);
                    ws.getRange(addr).setNumberValue(v);
                    srcTotal += v;
                } else {
                    ws.getRange(addr).setValue(srcData[ri][ci]);
                }
            }
        }
        System.out.println("[1] ソースデータ: " + srcData.length + " 行, 数量の合計 = " + srcTotal);

        // ---------- 2. ピボットテーブル作成 (行=地域, 値=数量 合計) ----------
        PivotTableWrapper pt = ws.getPivotTables().addPivotTable("売上集計", "A1:C11", "E1");

        PivotFieldObjectWrapper rf = pt.getField("地域");
        rf.setFieldType(XlPivotFieldType.PivotFieldTypeRow);
        PivotFieldObjectWrapper df = pt.getField("数量");
        df.setFieldType(XlPivotFieldType.PivotFieldTypeData);
        df.setDataFieldSubTotalsMethod(XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeSum);

        List<PivotFieldObjectWrapper> empty = new ArrayList<>();
        pt.setFields(pt.getPivotTableSetting(), XlPivotTableLayoutType.TabularForm,
                     Arrays.asList(rf), empty, empty, Arrays.asList(df));
        pt.refreshData();
        pt.refresh();
        System.out.println("[2] ピボットテーブル作成 (行=地域, 値=数量 合計, レイアウト=テーブル形式)");

        // ---------- 3. 集計結果の読み戻しと検算 ----------
        // ピボット出力: E1=ヘッダ, E2〜=行ラベル, F2〜=値, 最終行=総計
        Map<String, Double> pivotRows = new LinkedHashMap<>();
        for (int r = 2; r <= 10; r++) {
            String label = ws.getRange("E" + r + ":E" + r).getValue().get("E" + r);
            String val   = ws.getRange("F" + r + ":F" + r).getValue().get("F" + r);
            if (label != null && !label.isEmpty()) {
                pivotRows.put(label, Double.parseDouble(val));
                System.out.println("    " + label + " = " + val);
            }
        }
        double grandTotal = pivotRows.get("総計");
        // node の Number 表示（205）に合わせるため、整数なら long で表示する
        String gt = (grandTotal == Math.rint(grandTotal))
            ? String.valueOf((long) grandTotal) : String.valueOf(grandTotal);
        System.out.println("[3] 検算: 総計 " + gt + " === ソース合計 " + srcTotal
            + " → " + (grandTotal == srcTotal ? "OK" : "NG"));

        // ---------- 4. 集計方法を平均に変更 ----------
        df.setDataFieldSubTotalsMethod(XlPivotFieldSubtotalsMethodType.SubtotalsMethodTypeAverage);
        pt.setFields(pt.getPivotTableSetting(), XlPivotTableLayoutType.TabularForm,
                     Arrays.asList(rf), empty, empty, Arrays.asList(df));
        pt.refreshData();
        pt.refresh();
        System.out.println("[4] 集計方法を合計→平均に変更");
        for (int r = 2; r <= 10; r++) {
            String label = ws.getRange("E" + r + ":E" + r).getValue().get("E" + r);
            String val   = ws.getRange("F" + r + ":F" + r).getValue().get("F" + r);
            if (label != null && !label.isEmpty()) {
                System.out.println("    " + label + " = " + val);
            }
        }

        // ---------- 5. フィールド情報の取得 ----------
        List<PivotFieldObjectWrapper> fields = pt.getFields();
        StringBuilder caps = new StringBuilder();
        for (int i = 0; i < fields.size(); i++) {
            if (i > 0) caps.append(", ");
            caps.append(fields.get(i).getCaption());
        }
        System.out.println("[5] getFields: " + fields.size() + " 個 → " + caps);
        System.out.println("    getLayoutType = " + pt.getLayoutType().getCode()
            + " (TabularForm=" + XlPivotTableLayoutType.TabularForm.getCode() + ")");

        // ---------- 6. ピボットテーブルの削除 ----------
        ws.getPivotTables().removePivotTable("売上集計");
        try {
            ws.getPivotTables().getPivotTable("売上集計");
            System.out.println("[6] 削除失敗");
        } catch (Throwable t) {
            System.out.println("[6] removePivotTable → 削除確認OK");
        }

        wb.save();
        wb.close();
    }
}