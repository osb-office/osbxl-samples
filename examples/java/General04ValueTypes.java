// value-types.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import java.util.Map;

public class General04ValueTypes {
    static AppWrapper app;
    static WorkSheetWrapper ws;

    // JS の JSON.stringify 相当の引用（サンプル内に制御文字は無い前提）
    static String jsonq(String s) {
        return "\"" + s.replace("\\", "\\\\").replace("\"", "\\\"") + "\"";
    }

    // 整数値の double は int 表記にする（JS の出力と揃える）
    static String num(double d) {
        if (d == Math.rint(d) && !Double.isInfinite(d)) return String.valueOf((long) d);
        return String.valueOf(d);
    }

    static void show(String a1) {
        RangeWrapper r = ws.getRange(a1 + ":" + a1);
        Map<String, String> disp = r.getValue();
        Map<String, String> raw = r.getValue(true);          // true で生値
        Map<String, String> fmt = r.getNumberFormat();
        System.out.println("  " + a1 + ": 表示値=" + jsonq(disp.get(a1))
                + " 生値=" + jsonq(raw.get(a1))
                + " 書式=" + jsonq(fmt.get(a1)));
    }

    public static void main(String[] args) {
        app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("value-types-java.xlsx");
        ws = wb.openWorkSheet("Sheet1");

        // --- 日付・時刻 ---
        // DateTimeObjectWrapper は引数なしで生成し、setYMD() / setHMS() で設定します
        DateTimeObjectWrapper date = new DateTimeObjectWrapper();
        date.setYMD(2024, 1, 15);

        DateTimeObjectWrapper datetime = new DateTimeObjectWrapper();
        datetime.setYMD(2024, 1, 15);
        datetime.setHMS(13, 45, 30);

        DateTimeObjectWrapper time = new DateTimeObjectWrapper();
        time.setHMS(9, 5, 0);

        System.out.println("[日付・時刻]");
        ws.getRange("A1:A1").setDateValue(date);                       // 日付のみ
        ws.getRange("A2:A2").setDateValue(datetime);                   // 日付 + 時刻
        ws.getRange("A3:A3").setDateValue(time);                       // 時刻のみ
        ws.getRange("A4:A4").setDateStringValue("2024-01-15");         // 文字列から
        ws.getRange("A5:A5").setDateStringValue("令和6年1月15日");       // 和暦も認識されます
        ws.getRange("A6:A6").setDateValue(date, false, "yyyy-mm-dd");   // 表示書式を同時に指定
        ws.getRange("A7:A7").setDateValue(date, false, "ggge年m月d日"); // 和暦+引用符なし日本語
        for (String a : new String[]{"A1", "A2", "A3", "A4", "A5", "A6", "A7"}) show(a);

        // --- 数値・文字列 ---
        System.out.println("[数値・文字列]");
        ws.getRange("B1:B1").setNumberValue(1234567.891);
        ws.getRange("B2:B2").setNumberValue(1234567.891, false, "#,##0.00");
        ws.getRange("B3:B3").setNumberValue(0.25, false, "0.0%");
        ws.getRange("B4:B4").setValue("00123");       // 先頭 0 は文字列として保持される
        ws.getRange("B5:B5").setValue("2024/1/15");   // 汎用 setter は日付として自動認識
        ws.getRange("B6:B6").setValue("12345", true); // forceString=true で強制的に文字列扱い
        ws.getRange("B7:B7").setNumberValue(1234.5, false, "\"円\"#,##0");    // 引用符リテラル
        ws.getRange("B8:B8").setNumberValue(-1234.5, false, "#,##0;赤-#,##0"); // 裸の色トークン
        for (String b : new String[]{"B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"}) show(b);

        // --- 真偽値 ---
        System.out.println("[真偽値]");
        ws.getRange("C1:C1").setBooleanValue(true);
        ws.getRange("C2:C2").setBooleanValue(false);
        ws.getRange("C3:C3").setFormula("TRUE");      // 数式でも真偽値になる（"=" は付けない）
        for (String c : new String[]{"C1", "C2", "C3"}) show(c);

        // --- 日付のシリアル値 ---
        System.out.println("[シリアル値]");
        System.out.println("  2024/1/15 のシリアル値 = " + num(app.getNumericValue(date)));
        System.out.println("  1904 年形式の場合      = " + num(app.getNumericValue(date, true)));
        System.out.println("  wb.isDate1904()        = " + wb.isDate1904());

        wb.save();
        wb.close();
    }
}