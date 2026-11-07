// rows-cols.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.enums.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class General05RowsCols {
    static AppWrapper app;
    static WorkSheetWrapper ws;
    static final Pattern ADDR = Pattern.compile("^([A-Z]+)(\\d+)$");

    // 行 → 列 の順に並べて 1 行で表示する（· は空セル）
    static void dump(String label, String range) {
        Map<String, String> v = ws.getRange(range).getValue();
        List<String> keys = new ArrayList<>(v.keySet());
        keys.sort((a, b) -> {
            Matcher ma = ADDR.matcher(a);
            Matcher mb = ADDR.matcher(b);
            ma.matches();
            mb.matches();
            int cmp = Integer.compare(Integer.parseInt(ma.group(2)), Integer.parseInt(mb.group(2)));
            if (cmp != 0) return cmp;
            return Integer.compare(app.convertToColumnNumber(ma.group(1)),
                                   app.convertToColumnNumber(mb.group(1)));
        });
        StringBuilder sb = new StringBuilder();
        for (String k : keys) {
            if (sb.length() > 0) sb.append(" ");
            String val = v.get(k);
            sb.append(k).append("=").append(val == null || val.isEmpty() ? "·" : val);
        }
        System.out.println(label + " " + sb);
    }

    // A1:C5 を r{行}c{列} で埋める
    static void fill() {
        for (int r = 1; r <= 5; r++)
            for (int c = 1; c <= 3; c++)
                ws.getCells(r, c).setValue("r" + r + "c" + c);
    }

    public static void main(String[] args) {
        app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("rows-cols-java.xlsx");
        ws = wb.openWorkSheet("Sheet1");

        fill();
        dump("[初期]       A列:", "A1:A5");

        // ---------- 行 ----------
        ws.insertRow(3, 2);                    // 3 行目の「前」に 2 行挿入 → 既存データは下へずれる
        dump("[insertRow]  A列:", "A1:A7");
        ws.deleteRow(3, 2);                    // 3 行目から 2 行削除 → 下の行が上がる
        dump("[deleteRow]  A列:", "A1:A5");
        ws.clearRow(2, 1);                     // 2 行目をクリア（位置はずらさない）
        dump("[clearRow]   A列:", "A1:A5");
        ws.copyRow("Sheet1", 1, 1, 4);         // 自シートの 1 行目を 4 行目へ上書きコピー
        dump("[copyRow]    A列:", "A1:A5");

        fill();
        // ---------- 列 ----------
        ws.insertCol(2, 1);                    // B 列の前に 1 列挿入
        dump("[insertCol] 1行目:", "A1:D1");
        ws.deleteCol(2, 1);                    // 挿入した B 列を削除
        dump("[deleteCol] 1行目:", "A1:C1");
        ws.clearCol(3, 1);                     // C 列をクリア
        dump("[clearCol]  1行目:", "A1:C1");
        fill();
        ws.copyCol("Sheet1", 1, 2, 5);         // A,B 列を E,F 列へコピー
        dump("[copyCol]   1行目:", "A1:F1");

        // ---------- セル範囲 ----------
        ws.insertCell("B2:B3", XlInsertDirection.InsertDirectionDown);
        dump("[insertCell] B列:", "B1:B5");
        ws.deleteCell("B2:B3", XlInsertDirection.DeleteDirectionUp);
        dump("[deleteCell] B列:", "B1:B5");
        ws.clearCell("A1:B2");
        dump("[clearCell] A1:B2:", "A1:B2");
        ws.copyCell("Sheet1", "A3:B4", "F8");
        dump("[copyCell]  F8:G9:", "F8:G9");

        wb.save();
        wb.close();
    }
}