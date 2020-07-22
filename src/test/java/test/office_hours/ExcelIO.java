package test.office_hours;

import org.apache.poi.ss.usermodel.*;
import org.testng.annotations.Test;

import java.io.FileInputStream;

public class ExcelIO {

    @Test
    public void readExcelFile(){
        try(FileInputStream fileInputStream = new FileInputStream("VytrackTestUsers.xlsx")){
            //.xlsx
            //XSSFWorkbook - to create an object of .xlsx  file
            //HSSFWorkbook - to create an object of .xls  file
            //Workbook - is an interface, that is implemented by XSSFWorkbook and HSSFWorkbook
            Workbook workbook = WorkbookFactory.create(fileInputStream);
            //Object of Sheet used to represent excel file page
            //because, 1 excel document can have many spreadsheets
            Sheet sheet = workbook.getSheet("QA1-short");
            //read a first row
            Row row = sheet.getRow(0);
            //object of Cell class represents some cell. Row consists of cells
            //read a first cell
            Cell cell = row.getCell(0);
            //read a value from 1st row 1st column
            System.out.println(cell);
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
