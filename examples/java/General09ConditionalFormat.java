// conditional-format.java
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import com.osboffice.osbxl.enums.*;
import java.util.Map;

public class General09ConditionalFormat {
    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("condfmt-java.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");

        // ---------- データ表（node 版と同一） ----------
        String[] headers = {"月", "セル値ルール", "カラースケール", "データバー", "上位/平均"};
        for (int i = 0; i < headers.length; i++) {
            String col = "ABCDE".substring(i, i + 1);
            ws.getRange(col + "1:" + col + "1").setValue(headers[i]);
        }
        int[] sales = {45, 78, 23, 91, 56, 34, 87, 12, 65, 49};
        for (int i = 0; i < sales.length; i++) {
            ws.getRange("A" + (i + 2) + ":A" + (i + 2)).setNumberValue(i + 1);
            for (String col : new String[]{"B", "C", "D", "E"}) {
                ws.getRange(col + (i + 2) + ":" + col + (i + 2)).setNumberValue(sales[i]);
            }
        }

        FormatConditionsWrapper fc = ws.getFormatConditions();

        // ---------- (1) セル値ルール: 60 より大きい ----------
        CellValueObjectWrapper cellVal = new CellValueObjectWrapper();
        cellVal.setValueCondition(60, XlFormatConditionOperator.FormatConditionGreater);
        try {
            fc.setCellValueCondition(1, cellVal);
            System.out.println("[1] セル値ルールを登録できました");
        } catch (RuntimeException e) {
            System.out.println("[1] セル値ルール登録は例外: " + e.getMessage());
        }
        fc.apply();

        // ---------- ルール一覧の読み戻し ----------
        Map<Integer, Map<String, XlFormatConditionType>> types = fc.getConditionTypes();
        System.out.println("[2] ルール数 = " + types.size());

        // ---------- 削除系 ----------
        fc.removeCondition(4);
        System.out.println("[5] removeCondition(4) 後のルール数 = " + fc.getConditionTypes().size());
        fc.removeAllCondition();
        System.out.println("[6] removeAllCondition() 後のルール数 = " + fc.getConditionTypes().size());

        wb.save();
        wb.close();
        System.out.println("[done] condfmt-java.xlsx を保存しました");
    }
}