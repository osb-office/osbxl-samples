import com.osb_office.osbxl.App;
import com.osb_office.osbxl.WorkBook;
import com.osb_office.osbxl.WorkSheet;

import java.io.File;

public class QuickStart {
    public static void main(String[] args) {
        String filePath = "inventory.xlsx";
        File file = new File(filePath);
        if (file.exists()) {
            file.delete();
        }

        App app = new App();
        WorkBook wb = app.createWorkBook(filePath);

        try {
            WorkSheet ws = wb.addWorkSheet("Inventory");

            // Header
            ws.getRange("A1").setValue("SKU");
            ws.getRange("B1").setValue("Stock");
            ws.getRange("C1").setValue("Reorder Threshold");

            // Data
            ws.getRange("A2").setValue("A-1001");
            ws.getRange("B2").setNumberValue(145);
            ws.getRange("C2").setNumberValue(50);

            ws.getRange("A3").setValue("B-2002");
            ws.getRange("B3").setNumberValue(32);
            ws.getRange("C3").setNumberValue(40);

            wb.save();
            System.out.println("Created: " + filePath);
        } finally {
            wb.close();
        }
    }
}
