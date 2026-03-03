import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;

public class ExcelReader {
    public static void main(String[] args) throws Exception {
        String filePath = "ESPECIALIDADES DEL PERSONAL MILITAR SUBALTERNO (EN PROCESO).xlsx";
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            System.out.println("Hojas disponibles:");
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                System.out.println("  - " + sheet.getSheetName());
                
                // Mostrar primeras 10 filas
                for (int r = 0; r < Math.min(10, sheet.getLastRowNum() + 1); r++) {
                    Row row = sheet.getRow(r);
                    if (row != null) {
                        System.out.print("    Fila " + r + ": ");
                        for (int c = 0; c < Math.min(5, row.getLastCellNum()); c++) {
                            Cell cell = row.getCell(c);
                            if (cell != null) {
                                System.out.print("[" + cell.toString() + "] ");
                            }
                        }
                        System.out.println();
                    }
                }
                System.out.println();
            }
        }
    }
}
