package excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;

public class write {
    public static void main(String[] args) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Data");

        // Write sample data to the Excel sheet
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(1);
        cell.setCellValue("Sample Data");

        try (FileOutputStream fos = new FileOutputStream("C:\\Users\\DELL\\Desktop\\excel.xlsx")) {
            workbook.write(fos);
            System.out.println("Excel file written successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}