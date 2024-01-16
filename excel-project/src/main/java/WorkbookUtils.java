import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import static java.nio.file.Files.newInputStream;
import static java.util.Objects.requireNonNull;
import static org.apache.poi.ss.usermodel.CellType.STRING;

public class WorkbookUtils {

    public static final String SHEET_APPROVED_SHEET = "Approved Sheet";
    public static final String CELL_NO_STATUS = "no status";

    private WorkbookUtils() {
    }

    public static File getExcelFile() {
        for (File file : requireNonNull(new File(".").listFiles())) {
            if (file.getAbsolutePath().endsWith(".xlsx")) {
                return file.getAbsoluteFile();
            }
        }
        return null;
    }

    public static Workbook getWorkbook() {
        try (Workbook workbook = new XSSFWorkbook(newInputStream(requireNonNull(getExcelFile()).toPath()))) {
            return workbook;
        } catch (IOException e) {
            throw new RuntimeException("No excel file is found");
        }
    }

    public static void removeApprovedSheet(Workbook workbook) {
        try {
            int i = -1;
            for (Sheet sheet : workbook) {
                ++i;
                if (sheet.getSheetName().equals(SHEET_APPROVED_SHEET)) {
                    workbook.removeSheetAt(i);
                    break;
                }
            }
            System.out.println(SHEET_APPROVED_SHEET + " is successfully deleted");
        } catch (IllegalArgumentException exception) {
            System.out.println(SHEET_APPROVED_SHEET + " is not found.");
        }
    }

    public static Integer getStatusAddress(Sheet sheet) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == STRING &&
                        cell.getStringCellValue().trim().equalsIgnoreCase(CELL_NO_STATUS)) {
                    return cell.getAddress().getColumn();
                }
            }
        }
        return null;
    }

    public static void saveExcelFile(Workbook workbook) {
        try (FileOutputStream fileOutputStream = new FileOutputStream(requireNonNull(getExcelFile()))) {
            workbook.write(fileOutputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
