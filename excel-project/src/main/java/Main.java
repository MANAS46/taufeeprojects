import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

import static java.util.Comparator.reverseOrder;

public class Main {

    public static void main(String[] args) {
        Workbook workbook = WorkbookUtils.getWorkbook();
        Sheet sheet = workbook.getSheetAt(0);
        List<Integer> approvedRow = new ArrayList<>();
        Integer statusAddress = WorkbookUtils.getStatusAddress(sheet);
        if (statusAddress == null) {
            throw new RuntimeException(WorkbookUtils.CELL_NO_STATUS + " column is not found.");
        }
        for (Row row : sheet) {
            Cell cell = row.getCell(statusAddress);
            if (cell.getStringCellValue().trim().equalsIgnoreCase("approved")) {
                approvedRow.add(cell.getRowIndex());
            }
        }
        if (!approvedRow.isEmpty()) {
            WorkbookUtils.removeApprovedSheet(workbook);
            Sheet newSheet = workbook.createSheet(WorkbookUtils.SHEET_APPROVED_SHEET);
            createHeadingRow(sheet, newSheet);
            for (int i = 0; i < approvedRow.size(); ++i) {
                Row oldRow = sheet.getRow(approvedRow.get(i));
                Row newRow = newSheet.createRow(i + 1);
                int j = 0;
                for (Cell cell : oldRow) {
                    Cell newRowCell = newRow.createCell(j++);
                    if (cell.getCellType() == CellType.STRING) {
                        newRowCell.setCellValue(cell.getStringCellValue());
                    } else if (cell.getCellType() == CellType.BOOLEAN) {
                        newRowCell.setCellValue(cell.getBooleanCellValue());
                    } else if (cell.getCellType() == CellType.NUMERIC) {
                        newRowCell.setCellValue(cell.getNumericCellValue());
                    }
                }
            }
        }
        approvedRow.stream()
                .sorted(reverseOrder())
                .forEach(row -> sheet.removeRow(sheet.getRow(row)));
        WorkbookUtils.saveExcelFile(workbook);
    }

    private static void createHeadingRow(Sheet oldSheet, Sheet newSheet) {
        int i = 0;
        Row topRow = newSheet.createRow(0);
        for (Cell cell : oldSheet.getRow(0)) {
            topRow.createCell(i++).setCellValue(cell.getStringCellValue());
        }
    }
}
