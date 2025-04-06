import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class ExcelWriter {
    public static void main(String[] args) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Students");

        Object[][] data = {
            {"Name", "Grade"},
            {"Eya", 18},
            {"Ali", 15},
            {"Sami", 13}
        };

        int rowCount = 0;
        for (Object[] rowData : data) {
            Row row = sheet.createRow(rowCount++);
            int colCount = 0;
            for (Object field : rowData) {
                Cell cell = row.createCell(colCount++);
                if (field instanceof String)
                    cell.setCellValue((String) field);
                else if (field instanceof Integer)
                    cell.setCellValue((Integer) field);
            }
        }

        try (FileOutputStream fos = new FileOutputStream("output.xlsx")) {
            workbook.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
