import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.util.List;
import java.util.stream.Collectors;

import static java.util.stream.IntStream.range;

public class Main {

    public static void main(String[] args) {
        Workbook workbook = WorkbookUtils.getWorkbook();
        WorkbookUtils.SHEET_NAMES.forEach(sheetName -> {
            Sheet workingSheet = workbook.getSheet(sheetName);
            System.out.println();
            printHeading(workingSheet.getSheetName() + " is processing");
            List<Integer> approvedRow = WorkbookUtils.getApprovedStatusIndex(workingSheet);
            if (!approvedRow.isEmpty()) {
                String approvedSheetName = workingSheet.getSheetName() + "-" + WorkbookUtils.SHEET_APPROVED_SHEET;
                WorkbookUtils.removeApprovedSheet(workbook, approvedSheetName);
                Sheet newSheet = workbook.createSheet(approvedSheetName);
                WorkbookUtils.copyHeadingRow(workingSheet, newSheet);
                CellCopyPolicy cellCopyPolicy = new CellCopyPolicy().createBuilder().build();
                printHeading(String.format("Moving rows from sheet: %s, to sheet: %s", workingSheet.getSheetName(), newSheet.getSheetName()));
                range(0, approvedRow.size())
                        .boxed()
                        .forEach(i -> {
                            int rowIndex = approvedRow.get(i);
                            Row oldRow = workingSheet.getRow(rowIndex);
                            ((XSSFRow) newSheet.createRow(i + 1))
                                    .copyRowFrom(oldRow, cellCopyPolicy);
                            workingSheet.removeRow(oldRow);
                            System.out.printf("Row %s is moved\n", rowIndex);
                        });
            }
//        WorkbookUtils.deleteRows(workingSheet, approvedRow);
        });
        WorkbookUtils.saveExcelFile(workbook);
    }

    private static void printHeading(String heading) {
        String border = range(0, heading.length()).mapToObj(i -> "=").collect(Collectors.joining());
        System.out.println(border);
        System.out.println(heading);
        System.out.println(border);
    }
}
