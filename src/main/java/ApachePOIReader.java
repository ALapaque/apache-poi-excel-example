import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

public class ApachePOIReader {

    public void read() {
        try {
            String filePath = KeyboardReader.askForKeyboardInteraction("Dossier d'enregistrement à partir de ? ");
            String fileName = KeyboardReader.askForKeyboardInteraction("Nom du fichier ? ");
            String extension = KeyboardReader.askForKeyboardInteraction("extension du fichier ? ");

            if (!filePath.startsWith("/")) filePath = "/" + filePath;
            if (!filePath.endsWith("/")) filePath = filePath + "/";
            if (!extension.startsWith(".")) extension = "." + extension;


            FileInputStream excelFile = new FileInputStream(new File(filePath + fileName + extension));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datasheet = workbook.getSheetAt(0);
            Iterator<Row> rows = datasheet.rowIterator();


            while (rows.hasNext()) {
                Row row = rows.next();

                Iterator<Cell> cells = row.iterator();

                while (cells.hasNext()) {
                    Cell cell = cells.next();

                    switch (cell.getCellTypeEnum()) {
                        case NUMERIC:
                            System.out.println("-- " + cell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            System.out.println("-- " + cell.getBooleanCellValue());
                            break;
                        case STRING:
                            System.out.println("--" + cell.getStringCellValue());
                            break;
                        default:
                            System.out.println("Cette cellule est d'un type " + cell.getCellTypeEnum());
                    }
                }
            }

        } catch (FileNotFoundException e) {
            System.out.println("Fichier introuvable");
            e.printStackTrace();
        } catch (IOException e) {
            System.out.println("Une erreur est survenue");
            e.printStackTrace();
        }
    }
}
