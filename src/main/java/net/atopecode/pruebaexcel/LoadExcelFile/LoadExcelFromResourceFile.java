package net.atopecode.pruebaexcel.LoadExcelFile;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.function.Consumer;
import java.util.function.Function;

/**
 * Esta clase carga una hoja de un archivo excel que esté en la carpeta 'Resources' del proyecto.
 */
public class LoadExcelFromResourceFile{
    public enum ExcelFormat { XSSF, HSSF };
    
    private String resourceFileName;
    private InputStream inputStreamExcel;
    private Workbook wb;
    private ExcelFormat excelFormat;

    public LoadExcelFromResourceFile(String resourceFileName, ExcelFormat excelFormat) throws IOException {
        this.resourceFileName = resourceFileName;
        this.excelFormat = excelFormat;
        
        loadXSSFWorkbook(resourceFileName, excelFormat);
    }
    
    private void loadXSSFWorkbook(String resourceFileName, ExcelFormat excelFormat) throws IOException{
        inputStreamExcel = getClass().getClassLoader().getResourceAsStream(resourceFileName);
        
        switch (excelFormat){
            case XSSF:
                wb = new XSSFWorkbook(new BufferedInputStream(inputStreamExcel));
                break;
                
            case  HSSF:
                wb = new HSSFWorkbook(new BufferedInputStream(inputStreamExcel));
                break;
        }
    }

    /**
     * Este método recorre todas las filas de una hoja del Excel y ejecuta un método (Consumer) recibido como 
     * parámetro por cada fila.
     * @param sheetName
     * @param numExcelHeaders
     * @param consumerRow
     * @throws Exception
     */
    public void iterateSheet(String sheetName, Integer numExcelHeaders, Consumer<Row> consumerRow){
        if(consumerRow == null){
            return;
        }
        
        Sheet sheet = wb.getSheet(sheetName);
        Iterator<Row> itr = sheet.rowIterator();

        Row row = null;

        //Cabeceras del Excel:
        for(int cont = 0; cont < numExcelHeaders; cont++){
            if(itr.hasNext()){
                row = itr.next();
            }
        }
        
        while(itr.hasNext()){
            row = itr.next();
            consumerRow.accept(row);
        }
    }

    public <T> ArrayList<T> sheetToArrayList(String sheetName, Integer numExcelHeaders, Function<Row, T> functionRow) throws Exception{
        Sheet sheet = getSheet(sheetName);
        return sheetToArrayList(sheet, numExcelHeaders, functionRow);
    }


    public <T> ArrayList<T> sheetToArrayList(int sheetNumber, Integer numExcelHeaders, Function<Row, T> functionRow) throws Exception{
        Sheet sheet = getSheet(sheetNumber);
        return sheetToArrayList(sheet, numExcelHeaders, functionRow);
    }
    
    /**
     * Este método recorre todas las filas de una hoja Excel y ejecuta un método (Function) recibido como 
     * parámetro para devolver un objeto 'Entidad' por cada fila del Excel y devolver un 'ArrayList<T>' con todas
     * las filas del Excel convertidas a objetos 'Entidad'.
     * @param numExcelHeaders
     * @param functionRow
     * @param <T>
     * @return 
     */
    public <T> ArrayList<T> sheetToArrayList(Sheet sheet, Integer numExcelHeaders, Function<Row, T> functionRow) throws Exception{
        ArrayList<T> entityList = new ArrayList<T>();
        if(functionRow == null){
            return entityList;
        }
        
        Iterator<Row> itr = sheet.rowIterator();

        Row row = null;

        //Cabeceras del Excel:
        for(int cont = 0; cont < numExcelHeaders; cont++){
            if(itr.hasNext()){
                row = itr.next();
            }
        }

        T entity = null;
        while(itr.hasNext()){
            row = itr.next();
            entity = functionRow.apply(row);
            if(entity != null){
                entityList.add(entity);   
            }
        }
        
        return entityList;
    }
    
    public boolean close(){
        try{
            inputStreamExcel.close();
            wb.close();
            return true;
        }
        catch(Exception ex){
            return false;
        }
    }
    
    public Sheet getSheet(String sheetName){
        return wb.getSheet(sheetName);
    }
    
    public Sheet getSheet(int sheetNumber){
        return wb.getSheetAt(sheetNumber);
    }

    public Iterator<Row> getRowIterator(int sheetNumber){
        Iterator<Row> iterator = null;
        Sheet sheet = getSheet(sheetNumber);
        if(sheet != null){
            iterator = sheet.rowIterator();
        }

        return iterator;
    }

    public Workbook getWorkBook(){
        return wb;
    }
}