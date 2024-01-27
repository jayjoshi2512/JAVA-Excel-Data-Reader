package ExcelReader;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
    public static void main(String[] args) {
        try (FileInputStream file = new FileInputStream("C:\\Users\\joshi\\OneDrive\\Desktop\\JAVA\\ExcelReader\\new.xlsx");
             XSSFWorkbook workbook = new XSSFWorkbook(file)) {

            XSSFSheet sheet = workbook.getSheetAt(0);
            int rows = sheet.getPhysicalNumberOfRows();
            int cols = sheet.getRow(0).getPhysicalNumberOfCells();

            String[][] matrix = new String[rows][cols];

            for (int i = 0; i < rows; i++) {
                Row row = sheet.getRow(i);
                for (int j = 0; j < cols; j++) {
                    Cell cell = row.getCell(j);
                    matrix[i][j] = cell.toString();
                }
            }

            // Print the matrix in tabular format
            printTable(matrix);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void printTable(String[][] matrix) {
        int[] maxColumnWidths = new int[matrix[0].length];

        // Find the maximum width of each column
        for (String[] row : matrix) {
            for (int j = 0; j < row.length; j++) {
                maxColumnWidths[j] = Math.max(maxColumnWidths[j], row[j].length());
            }
        }

        // Print the upper line
        printLine(maxColumnWidths);

        // Print the matrix in tabular format
        for (int i = 0; i < matrix.length; i++) {
            System.out.print("|");
            for (int j = 0; j < matrix[i].length; j++) {
                System.out.printf(" %-" + (maxColumnWidths[j] + 2) + "s |", matrix[i][j]);
            }
            System.out.println();

            // Print the separator line after the first row
            if (i == 0) {
                printLine(maxColumnWidths);
            }
        }

        // Print the bottom line
        printLine(maxColumnWidths);
    }

    private static void printLine(int[] columnWidths) {
        for (int width : columnWidths) {
            System.out.print("+" + "-".repeat(width + 4));
        }
        System.out.println("+");
    }
}