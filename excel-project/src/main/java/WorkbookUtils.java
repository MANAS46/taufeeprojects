import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

import static java.nio.file.Files.newInputStream;
import static java.util.Objects.requireNonNull;
import static org.apache.commons.collections4.list.UnmodifiableList.unmodifiableList;
import static org.apache.poi.ss.usermodel.CellType.STRING;

public class WorkbookUtils {

    private static final String STATUS_APPROVED = "approved";
    public static final String SHEET_APPROVED_SHEET = "Approved Sheet";
    public static final String CELL_NO_STATUS = "no status";
    public static final List<String> SHEET_NAMES = unmodifiableList(Arrays.asList("01 Lot", "0130 Lot"));

    private WorkbookUtils() {
    }

    public static File getExcelFile() {
        File rootFolder = new File(".");
        for (File file : requireNonNull(rootFolder.listFiles())) {
            if (file.getAbsolutePath().endsWith(".xlsx")) {
                return file.getAbsoluteFile();
            }
        }
        throw new RuntimeException("No .xlsx file is found in this path: " + rootFolder.getAbsolutePath());
    }

    public static Workbook getWorkbook() {
        try {
            return new XSSFWorkbook(newInputStream(getExcelFile().toPath()));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void removeApprovedSheet(Workbook workbook, String sheetName) {
        try {
            int i = -1;
            for (Sheet sheet : workbook) {
                ++i;
                if (sheet.getSheetName().equals(sheetName)) {
                    workbook.removeSheetAt(i);
                    System.out.println(sheetName + " is successfully deleted");
                    break;
                }
            }
            throw new IllegalArgumentException("");
        } catch (IllegalArgumentException exception) {
            System.out.println(sheetName + " is not found");
        }
    }

    public static List<Integer> getApprovedStatusIndex(Sheet sheet) {
        int statusAddress = getStatusAddress(sheet);
        return StreamSupport.stream(sheet.spliterator(), false)
                .map(row -> row.getCell(statusAddress))
                .filter(Objects::nonNull)
                .filter(cell -> cell.getStringCellValue().trim().equalsIgnoreCase(STATUS_APPROVED))
                .map(Cell::getRowIndex)
                .collect(Collectors.toList());
    }

    public static void copyHeadingRow(Sheet oldSheet, Sheet newSheet) {
        ((XSSFRow) newSheet.createRow(0))
                .copyRowFrom(oldSheet.getRow(0),
                        new CellCopyPolicy().createBuilder().build());
    }

    public static void deleteRows(Sheet sheet, List<Integer> approvedRow) {
        System.out.println("Deleting the rows");
        int[] indexes = StreamSupport.stream(sheet.spliterator(), false)
                .mapToInt(Row::getRowNum)
                .filter(index -> index > 0)
                .toArray();
        int start = 0, current = start, rowsShift = start;
        while (current < indexes.length) {
            if (indexes[current] + 1 == indexes[current + 1]) {
                ++current;
                continue;
            }
            if (indexes[start] - 1 != start) {
                sheet.shiftRows(indexes[start], indexes[current], -(rowsShift + 1), true, true);
            }
            rowsShift += current - start;
            ++current;
            start = current;
        }
        System.out.println("Rows are deleted");
    }

    public static void saveExcelFile(Workbook workbook) {
        AtomicBoolean threadRun = new AtomicBoolean(true);
        Thread timer = new Thread(() -> {
            LocalDateTime time = LocalDateTime.now();
            while (threadRun.get()) {
                if (time.plusSeconds(3).isBefore(LocalDateTime.now())) {
                    System.out.print(".");
                    time = LocalDateTime.now();
                }
            }
            System.out.println();
        });
        File file = getExcelFile();
        System.out.print(file.getName() + " is saving, please wait");
        try (FileOutputStream fileOutputStream = new FileOutputStream(file)) {
            timer.start();
            workbook.write(fileOutputStream);
            workbook.close();
            threadRun.set(false);
            Thread.sleep(1500);
            System.out.println("File saved successfully");
        } catch (IOException | InterruptedException e) {
            timer.interrupt();
            throw new RuntimeException(e);
        }
    }

    private static Integer getStatusAddress(Sheet sheet) {
        for (Cell cell : sheet.getRow(0)) {
            if (cell.getCellType() == STRING &&
                    cell.getStringCellValue().trim().equalsIgnoreCase(CELL_NO_STATUS)) {
                return cell.getAddress().getColumn();
            }
        }
        throw new RuntimeException(WorkbookUtils.CELL_NO_STATUS + " column is not found");
    }

}
