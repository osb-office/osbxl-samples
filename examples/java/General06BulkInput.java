// bulk-input.java — 2次元配列による値と数式の一括高速投入 (javaosbxl)
import com.osboffice.osbxl.*;
import java.util.*;

public class General06BulkInput {
    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("bulk-input.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");

        // ---------- 1. 2次元配列による値の一括高速投入 (setRangeData) ----------
        List<List<String>> data = Arrays.asList(
            Arrays.asList("商品名", "単価", "数量", "金額"),
            Arrays.asList("りんご", "100", "3"),
            Arrays.asList("みかん", "50", "6"),
            Arrays.asList("ぶどう", "150", "2")
        );
        ws.setRangeData("A1", data);

        System.out.println("[1] setRangeData 投入後:");
        for (String a1 : Arrays.asList("A1", "B2", "C2", "D1")) {
            System.out.println("    " + a1 + " = " + ws.getRange(a1).getValue().get(a1));
        }

        // ---------- 2. 2次元配列による数式の一括高速投入 (setRangeFormulas) ----------
        List<List<String>> formulas = Arrays.asList(
            Collections.singletonList("B2*C2"),
            Collections.singletonList("B3*C3"),
            Collections.singletonList("B4*C4")
        );
        ws.setRangeFormulas("D2", formulas);

        System.out.println("[2] setRangeFormulas 投入後:");
        System.out.println("    数式 D2:D4 = " + ws.getRange("D2:D4").getFormula());
        System.out.println("    計算値 D2:D4 = " + ws.getRange("D2:D4").getValue(true));

        // ---------- 3. 速度比較（2,000 行 x 5 列 = 10,000 セル）----------
        final int ROWS = 2000, COLS = 5;

        // セルごとに getCells() して投入
        WorkSheetWrapper wsCell = wb.addWorkSheet("CellByCell");
        long t = System.currentTimeMillis();
        for (int r = 1; r <= ROWS; r++) {
            for (int c = 1; c <= COLS; c++) wsCell.getCells(r, c).setNumberValue(r * 100 + c);
        }
        long msCell = System.currentTimeMillis() - t;

        // 2次元配列で setRangeData() 投入 (最高速)
        WorkSheetWrapper wsBulk = wb.addWorkSheet("Bulk");
        t = System.currentTimeMillis();
        List<List<String>> matrix = new ArrayList<>();
        for (int r = 1; r <= ROWS; r++) {
            List<String> row = new ArrayList<>();
            for (int c = 1; c <= COLS; c++) row.add(String.valueOf(r * 100 + c));
            matrix.add(row);
        }
        long msBuild = System.currentTimeMillis() - t;
        t = System.currentTimeMillis();
        wsBulk.setRangeData("A1", matrix);
        long msApply = System.currentTimeMillis() - t;

        System.out.println("[3] " + (ROWS * COLS) + " セルの投入速度比較:");
        System.out.println("    セルごと getCells().setNumberValue() : " + msCell + " ms");
        System.out.println("    setRangeData (構築" + msBuild + " + 投入" + msApply + ") : " + (msBuild + msApply) + " ms (最速)");

        String a = wsCell.getRange("E2000:E2000").getValue(true).get("E2000");
        String b = wsBulk.getRange("E2000:E2000").getValue(true).get("E2000");
        System.out.println("[検証] 両方式とも E2000 = " + (a.equals(b) ? b : "不一致 " + a + " / " + b));

        wb.save();
        wb.close();
    }
}