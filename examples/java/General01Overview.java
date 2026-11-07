// basic-flow.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import java.util.Map;

public class General01Overview {
    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();                          // [1] App を生成

        WorkBookWrapper wb = app.createWorkBook("flow-java.xlsx");  // [2] 新規作成
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");           // [3] ワークシートを開く

        ws.getRange("A1:A1").setValue("新規作成");
        ws.getRange("B1:B1").setFormula("100*3");                   // [5] 数式（先頭の "=" は付けない）
        Map<String, String> v = ws.getRange("B1:B1").getValue(true);
        System.out.println("[6] getValue(true) -> B1 = " + v.get("B1"));

        wb.save();                                                  // [7] 保存
        wb.close();                                                 // [8] 閉じる
        System.out.println("[8] close() 完了");

        wb = app.openWorkBook("flow-java.xlsx");                    // [2'] 既存のブックを開く
        ws = wb.openWorkSheet("Sheet1");
        Map<String, String> r = ws.getRange("A1:A1").getValue(true);
        System.out.println("[2'] openWorkBook -> A1 = " + r.get("A1"));
        wb.close();
        System.out.println("[8] close() 完了");
    }
}