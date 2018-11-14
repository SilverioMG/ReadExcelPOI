package net.atopecode.pruebaexcel;

import net.atopecode.pruebaexcel.LoadExcelFile.LoadExcelFromResourceFile;
import net.atopecode.pruebaexcel.LoadExcelFile.ResourcesFile;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.util.Base64;
import java.util.Iterator;

public class Application {

    public static void main(String[] args){
        String filePath = "Archivo1.xlsx";

        //Carga del archivo excel y lectura.
        try{
            LoadExcelFromResourceFile loadExcel = new LoadExcelFromResourceFile(filePath, LoadExcelFromResourceFile.ExcelFormat.XSSF);
            Iterator<Row> it = loadExcel.getRowIterator(0);
            while(it.hasNext()){
                Row row = it.next();
                System.out.print(row.toString());
            }
        }
        catch(Exception ex){
            System.out.print("ERROR - " + ex.getMessage());
        }

        //Carga del archivo en la carpeta 'resources' del proyecto en 'byte[]' y generación del excel en memoria.
        try{
            ResourcesFile resourcesFile = new ResourcesFile(filePath);
            String base64File = resourcesFile.loadResourceFileBase64();
            byte[] bytes = Base64.getDecoder().decode(base64File); //resourcesFile.loadResourceFileBytes();
            XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes));
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> it = sheet.rowIterator();
            while(it.hasNext()){
                Row row = it.next();
                System.out.print(row.toString());
            }
        }
        catch(Exception ex){
            System.out.print("ERROR - " + ex.getMessage());
        }
    }
}
