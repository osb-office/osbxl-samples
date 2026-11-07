// sheets.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class General03Sheets {
    static WorkBookWrapper wb;

    // インデックスは 0 始まり (zeroBase)。0..getSheetCount()-1 を順に開いて名前を並べる
    static void show(String label) {
        List<String> names = new ArrayList<>();
        for (int i = 0; i < wb.getSheetCount(); i++) {
            names.add(wb.openWorkSheet(i).getName());
        }
        System.out.println(label + " count=" + wb.getSheetCount()
                + " [" + String.join(", ", names) + "]");
    }

    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        wb = app.createWorkBook("sheets-java.xlsx");

        show("[1] 新規作成直後        :");

        wb.addWorkSheet("集計");                       // position 省略 → 末尾に追加
        show("[2] addWorkSheet(集計)  :");

        wb.addWorkSheet("受注", 0);                    // position=0 → 先頭に挿入 (0-based)
        show("[3] addWorkSheet(受注,0):");

        wb.moveSheet("受注", 2);                       // 3 番目へ移動 (0-based: index 2)
        show("[4] moveSheet(受注,2)   :");

        WorkSheetWrapper ws = wb.openWorkSheet("受注");
        System.out.println("[5] getName() = " + ws.getName()
                + " / getSheetNumber(受注) = " + wb.getSheetNumber("受注"));
        ws.getRange("A1:A1").setValue("受注データ");

        wb.openWorkSheet("Sheet1").setName("明細");    // リネーム
        show("[6] setName(明細)       :");

        // --- アクティブタブ（ブックを開いたときに手前に表示されるシート）---
        System.out.println("[7] isActive: 集計 = " + wb.openWorkSheet("集計").isActive()
                + " / 明細 = " + wb.openWorkSheet("明細").isActive());
        wb.openWorkSheet("集計").activate(true);       // 他のシートのアクティブは自動的に解除されます
        System.out.println("    activate後: 集計 = " + wb.openWorkSheet("集計").isActive()
                + " / 明細 = " + wb.openWorkSheet("明細").isActive());

        // --- 削除 ---
        wb.addWorkSheet("一時");
        wb.deleteWorkSheet("一時");
        show("[8] deleteWorkSheet(一時):");

        wb.save();
        wb.close();

        // --- 開き直して確認 ---
        WorkBookWrapper wb2 = app.openWorkBook("sheets-java.xlsx");
        List<String> names = new ArrayList<>();
        for (int i = 0; i < wb2.getSheetCount(); i++) names.add(wb2.openWorkSheet(i).getName());
        System.out.println("[終] シート構成: " + String.join(", ", names));
        Map<String, String> v = wb2.openWorkSheet("受注").getRange("A1:A1").getValue();
        System.out.println("     受注!A1 = " + v.get("A1"));
        wb2.close();
    }
}