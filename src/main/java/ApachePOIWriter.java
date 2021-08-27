import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class ApachePOIWriter {

    public void write() {
        // create a workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        // create a default sheet named First sheet
        XSSFSheet sheet = workbook.createSheet("First sheet");
        // contains an object which can be iterate through columns and rows
        Object[][] dataObject = {
                {"title"},
                {"header1", "header2", "header3"},
                {"col1", "col2", "col3"}
        };

        int rowNumber = 0;
        System.out.println("Create the excel file");

        for (Object[] data : dataObject) {
            // create a row for each row wanted
            Row row = sheet.createRow(rowNumber++);

            int columnNumber = 0;

            for (Object field : data) {
                // create a cell for each cell wanted
                Cell cell = row.createCell(columnNumber++);

                if (rowNumber == 1 && columnNumber == 1) {
                    sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, data.length + 1));
                }

                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else {
                    cell.setCellValue("--ERROR--");
                }
            }
        }


        try {
            String filePath = KeyboardReader.askForKeyboardInteraction("Dossier d'enregistrement à partir de ? ");
            if (!filePath.startsWith("/")) filePath = "/" + filePath;
            if (!filePath.endsWith("/")) filePath = filePath + "/";

            Path path = Paths.get(filePath);

            Files.createDirectories(path);

            String fileName = KeyboardReader.askForKeyboardInteraction("Nom du fichier ? ");
            String extension = KeyboardReader.askForKeyboardInteraction("extension du fichier ? ");

            if (!extension.startsWith(".")) extension = "." + extension;

            FileOutputStream stream = new FileOutputStream(filePath + fileName + extension);

            workbook.write(stream);
            System.out.println("Fichier créé avec succès");
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Une erreur est survenue");
        }
    }
}
