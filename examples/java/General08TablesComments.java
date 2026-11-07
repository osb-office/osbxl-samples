// tables-comments.java — node 版サンプルと同一検証 (javaosbxl)
import com.osboffice.osbxl.*;
import com.osboffice.osbxl.dto.*;
import com.osboffice.osbxl.enums.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

public class General08TablesComments {
    // JSON.stringify 相当の出力ヘルパー（Map は挿入順のまま）
    static String jsonMap(Map<String, String> m) {
        StringBuilder sb = new StringBuilder("{");
        boolean first = true;
        for (Map.Entry<String, String> e : m.entrySet()) {
            if (!first) sb.append(",");
            first = false;
            sb.append("\"").append(e.getKey()).append("\":\"").append(e.getValue()).append("\"");
        }
        return sb.append("}").toString();
    }

    static String jsonList(List<String> l) {
        StringBuilder sb = new StringBuilder("[");
        for (int i = 0; i < l.size(); i++) {
            if (i > 0) sb.append(",");
            sb.append("\"").append(l.get(i)).append("\"");
        }
        return sb.append("]").toString();
    }

    static CommentObjectWrapper mkComment(String author, String content) {
        CommentObjectWrapper c = new CommentObjectWrapper();
        c.setAuthor(author);
        c.setContent(content);
        return c;
    }

    public static void main(String[] args) {
        AppWrapper app = new AppWrapper();
        WorkBookWrapper wb = app.createWorkBook("tables-comments-java.xlsx");
        WorkSheetWrapper ws = wb.openWorkSheet("Sheet1");

        // ---------- データ投入（ヘッダ1行 + データ3行）----------
        List<List<String>> rows = Arrays.asList(
            Arrays.asList("商品名", "単価", "数量"),
            Arrays.asList("リンゴ", "120", "3"),
            Arrays.asList("ミカン", "80", "5"),
            Arrays.asList("ブドウ", "300", "2")
        );
        ws.setRangeData("A1", rows);

        // ---------- 1. テーブルを作成 ----------
        ListObjectsWrapper lo = ws.getListObjects();
        TableWrapper table = lo.addList("Sales", "A1:C4", true, true); // useFirstRowAsHeader, insertTotals
        System.out.println("[1] テーブル作成");
        System.out.println("    名前      = " + table.getName());
        System.out.println("    列名      = " + jsonList(table.getColumnName()));
        System.out.println("    スタイル  = " + table.getTableStyleName());
        System.out.println("    全範囲    = " + table.getAllRange().getAddress());
        System.out.println("    ヘッダ    = " + table.getHeaderRowRange().getAddress());
        System.out.println("    データ部  = " + table.getDataBodyRange().getAddress());
        System.out.println("    集計行    = " + table.getTotalRowRange().getAddress());

        // ---------- 2. 集計行 ----------
        table.setTotalRowLabel(0, "合計");
        table.setTotalRowFunction(1, XlTotalsCalculation.TotalsCalculationSum);
        table.setTotalRowFunction(2, XlTotalsCalculation.TotalsCalculationSum);
        System.out.println("[2] 集計行");
        System.out.println("    ラベル(1列目 (0-based: 0)) = \"" + table.getTotalRowLabel(1) + "\"");
        System.out.println("    数式          = " + jsonMap(ws.getRange("B5:C5").getFormula()));
        System.out.println("    値            = " + jsonMap(ws.getRange("B5:C5").getValue(true)));

        // ---------- 3. テーブルスタイルを変更 ----------
        table.setBuiltinStyleName(XlDefaultTableStyle.TableStyleMedium9, true);
        System.out.println("[3] スタイル変更後 = " + table.getTableStyleName());

        // ---------- 4. 構造化参照（テーブル列を参照する数式）----------
        ws.getRange("E1:E1").setFormula("SUM(Sales[単価])"); // 先頭に = は付けない
        System.out.println("[4] 構造化参照");
        System.out.println("    E1 数式 = " + jsonMap(ws.getRange("E1:E1").getFormula()));
        System.out.println("    E1 値   = " + jsonMap(ws.getRange("E1:E1").getValue(true)));

        // ---------- 5. コメント（新式）とメモ（旧式）----------
        CommentsWrapper cm = ws.getComments();
        // 新式コメント（スレッド）: 本体 + 返信の 2 件 → B2（save 後もファイルに残す）
        cm.setCommentThread("B2", Arrays.asList(
            mkComment("alice", "この数値を確認"), mkComment("bob", "確認しました")));
        // 単一の新式コメント → D2（削除の実演用）
        cm.setComment("D2", mkComment("dave", "あとで消すコメント"));
        // 旧式メモ（吹き出し）→ A2
        cm.setMemo("A2", "Carol", "最優先商品です", new FontObjectWrapper());

        List<CommentObjectWrapper> thread = cm.getComment("B2");
        System.out.println("[5] コメント");
        System.out.println("    B2 スレッド件数 = " + thread.size());
        for (int i = 0; i < thread.size(); i++) {
            CommentObjectWrapper c = thread.get(i);
            String pid = c.getParentId();
            boolean isReply = pid != null && !pid.isEmpty();
            System.out.println("      [" + i + "] " + c.getAuthor() + ": " + c.getContent()
                + (isReply ? " (返信)" : ""));
        }
        System.out.println("    D2 コメント件数 = " + cm.getComment("D2").size());
        System.out.println("    A2 メモ = \"" + cm.getMemoText("A2") + "\" / 作者 \"" + cm.getMemoAuthor("A2") + "\"");

        // 削除: 新式は removeComment / 旧式は removeMemo（削除後の読み取りは "not found" 例外）
        cm.removeComment("D2");
        cm.removeMemo("A2");
        try {
            cm.getComment("D2");
            System.out.println("    削除後 D2コメント: 例外なし");
        } catch (RuntimeException e) {
            System.out.println("    削除後 D2コメント: 例外 " + e.getMessage());
        }
        try {
            cm.getMemoText("A2");
            System.out.println("    削除後 A2メモ: 例外なし");
        } catch (RuntimeException e) {
            System.out.println("    削除後 A2メモ: 例外 " + e.getMessage());
        }
        System.out.println("    B2 は残すので save 後も " + cm.getComment("B2").size() + "件");

        String total = ws.getRange("B5:B5").getValue(true).get("B5");
        String e1 = ws.getRange("E1:E1").getValue(true).get("E1");
        System.out.println("[確認] テーブル=" + table.getName() + " 集計(単価)=" + total + " 構造化参照(E1)=" + e1);

        wb.save();
        wb.close();
    }
}