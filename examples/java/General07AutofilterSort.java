// af-sort.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import com.osboffice.osbxl.enums.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class General07AutofilterSort {
    static WorkSheetWrapper ws;

    // 表示用ヘルパー: 行番号 → "商品名(値)" / 非表示でない行の一覧
    static String text(int r, String col) {
        return ws.getRange(col + r + ":" + col + r).getValue().get(col + r);
    }

    static String label(int r, String col) {
        return text(r, "A") + "(" + text(r, col) + ")";
    }

    static String visible(String col) {
        List<String> out = new ArrayList<>();
        for (int r = 2; r <= 11; r++) {
            if (!ws.getRow(r).getHidden()) out.add(label(r, col));
        }
        return String.join(" ", out);
    }

    static String hiddenList() {
        List<String> out = new ArrayList<>();
        for (int r = 2; r <= 11; r++) {
            if (ws.getRow(r).getHidden()) out.add(String.valueOf(r));
        }
        return out.isEmpty() ? "なし" : String.join(",", out);
    }

    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("af-sort-java.xlsx");
        ws = wb.openWorkSheet("Sheet1");

        // ---------- 1. データ準備（ヘッダ + 10 行）----------
        Object[][] DATA = {
            {"りんご", "東京", 120.0, 100.0},
            {"みかん", "大阪", 40.0, 80.0},
            {"バナナ", "東京", 200.0, 150.0},
            {"ぶどう", "福岡", 60.0, 300.0},
            {"なし", "大阪", 90.0, 120.0},
            {"もも", "福岡", 150.0, 400.0},
            {"キウイ", "東京", 30.0, 90.0},
            {"マンゴー", "大阪", 70.0, 500.0},
            {"いちご", "福岡", 110.0, 350.0},
            {"メロン", "東京", 25.0, 800.0},
        };
        String[] headers = {"商品名", "地域", "数量", "単価"};
        for (int i = 0; i < headers.length; i++) {
            String col = String.valueOf("ABCD".charAt(i));
            ws.getRange(col + "1:" + col + "1").setValue(headers[i]);
        }
        for (int i = 0; i < DATA.length; i++) {
            int r = i + 2;
            ws.getRange("A" + r + ":A" + r).setValue((String) DATA[i][0]);
            ws.getRange("B" + r + ":B" + r).setValue((String) DATA[i][1]);
            ws.getRange("C" + r + ":C" + r).setNumberValue((Double) DATA[i][2]);
            ws.getRange("D" + r + ":D" + r).setNumberValue((Double) DATA[i][3]);
        }
        System.out.println("[1] データ準備: A1:D11（ヘッダ + 10 行）");

        // ---------- 2. オートフィルター設定 ----------
        AutoFilterWrapper af = ws.getAutoFilter("A1:D11", true);   // 先頭行はヘッダ
        System.out.println("[2] getAutoFilter → 範囲 = " + af.getAddress() + " / 非表示行 = " + hiddenList());

        // ---------- 3. 条件フィルター: 数量 > 100 ----------
        // colId は範囲内の列番号（A=1 … D=4）。criteria に演算子記号は含めない
        af.setCustomFilter(XlAutoFilterOperator.GreaterThan, "100", 3);
        System.out.println("[3] 数量>100 の可視行: " + visible("C"));
        System.out.println("    非表示行 = " + hiddenList()
                + " / getValue(\"A2:A11\") の件数 = " + ws.getRange("A2:A11").getValue().size());
        af.resetAllFilter();                            // 条件だけ解除（フィルター自体は残る）

        // ---------- 4. Top-N / 値リストのフィルター ----------
        af.setTop10ValueFilter(3, 3);                   // 数量の上位 3 件
        System.out.println("[4] 数量上位3件の可視行: " + visible("C"));
        af.resetAllFilter();
        af.setUniqueValuesNumberFilter(Arrays.asList(120.0, 40.0), 3);   // 数量が 120 または 40
        System.out.println("[5] 数量∈{120,40} の可視行: " + visible("C"));
        af.resetAllFilter();
        af.setUniqueValuesStringFilter(Arrays.asList("東京", "福岡"), 2);   // 地域が 東京 または 福岡（文字列リスト）
        System.out.println("[6] 地域∈{東京,福岡} の可視行: " + visible("B"));
        af.resetAllFilter();
        af.setCustomFilter(XlAutoFilterOperator.Equal, "大阪", 2);   // 地域 = 大阪（文字列一致）
        System.out.println("[7] 地域=\"大阪\" の可視行: " + visible("B"));
        af.resetAllFilter();

        // ---------- 5. 並べ替え（単一キー）----------
        // ヘッダを含まないデータ部だけを範囲にする（firstRowAsHeader 省略 = false）
        ws.getSort("A2:D11").executeSortAscending(2);   // 数量（範囲内3列目 (0-based: 2)）で昇順
        System.out.println("[8] 数量の昇順: " + visible("C"));

        // オートフィルター経由の Sort はヘッダ行が自動的に保護される
        SortWrapper sort = af.getSort();
        sort.executeSortDescending(3);                  // 単価（4列目 (0-based: 3)）で降順
        System.out.println("[9] 単価の降順: " + visible("D"));

        // ---------- 6. 並び替え条件の読み戻し ----------
        SortFieldObjectWrapper cond = sort.getSortConditions().get(0);
        System.out.println("[10] 条件の読み戻し: 対象列 = " + cond.getTarget()
                + " / 昇順 = " + cond.getSortAscending());

        // ---------- 7. オートフィルター自体の削除 ----------
        af.removeFilter();
        System.out.println("[11] removeFilter 後: 非表示行 = " + hiddenList());

        wb.save();
        wb.close();
    }
}