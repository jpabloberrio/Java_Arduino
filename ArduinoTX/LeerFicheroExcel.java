package ArduinoTX;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.Vector;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LeerFicheroExcel {
    public  Vector<String> Preguntas = new Vector<String>();
    
    public  void LeerFicheroExcel() {
        String nombreArchivo = "test.xlsx";
        String rutaArchivo = "C:\\nuevo\\" + nombreArchivo;
        String hoja = "Hoja1";

        try (FileInputStream file = new FileInputStream(new File(rutaArchivo))) {
            // leer archivo excel
            XSSFWorkbook worbook = new XSSFWorkbook(file);
            //obtener la hoja que se va leer
            XSSFSheet sheet = worbook.getSheetAt(0);
            //obtener todas las filas de la hoja excel
            Iterator<Row> rowIterator = sheet.iterator();
            //String[] Preguntas;
            
            Row row;
            // se recorre cada fila hasta el final
            while (rowIterator.hasNext()) {
                row = rowIterator.next();
                //se obtiene las celdas por fila
                Iterator<Cell> cellIterator = row.cellIterator();
                Cell cell;
                //se recorre cada celda
                while (cellIterator.hasNext()) {
                    // se obtiene la celda en específico y se la imprime
                       
                    cell = cellIterator.next();
                    Preguntas.add(cell.getStringCellValue());
                    System.out.print(cell.getStringCellValue() + " | ");
                }
                System.out.println();
            }
        } catch (Exception e) {
            e.getMessage();
        }
        
        System.out.print(Preguntas.elementAt(22));
    }
}
