package dgsam;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;

public class ExcelAnalyzer {
    public static void main(String[] args) throws Exception {
        String filePath = "ESPECIALIDADES DEL PERSONAL MILITAR SUBALTERNO (EN PROCESO).xlsx";
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            System.out.println("=== Hojas disponibles ===");
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                System.out.println("Hoja " + i + ": " + sheet.getSheetName());
                
                // Mostrar primeras 15 filas
                System.out.println("  Primeras filas:");
                for (int r = 0; r < Math.min(15, sheet.getLastRowNum() + 1); r++) {
                    Row row = sheet.getRow(r);
                    if (row != null) {
                        StringBuilder sb = new StringBuilder("    Fila " + r + ": ");
                        for (int c = 0; c < row.getLastCellNum(); c++) {
                            Cell cell = row.getCell(c);
                            String val = cell == null ? "" : cell.toString().trim();
                            if (!val.isEmpty()) {
                                sb.append("[Col").append(c).append("=").append(val).append("] ");
                            }
                        }
                        System.out.println(sb.toString());
                    }
                }
                System.out.println();
            }
        }
    }
}
