package com.belwar.aiseexcelprocessor;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Scanner;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelChecker {

    private Scanner scanner;
    
    public ExcelChecker (Scanner scanner) {
    
        this.scanner = scanner;
        //Scanner scanner;
        String folder = "", group = "";

        //scanner = new Scanner(System.in);
        
        System.out.println("Choose the target folder: ");
        folder = scanner.nextLine();

        System.out.println("Choose the group number/name: ");
        group = scanner.nextLine();
        
        //scanner.close();
        
        System.out.println("** Reading folder " + folder  +  " **");

	    File dir = new File(folder);
        File[] filesOnDir = dir.listFiles();
	    
	    System.out.println("** " + filesOnDir.length  +  " files found **");
	    
        for (File f : filesOnDir) {
            if (f.isFile() && f.getName().endsWith(".xlsx"))
                CheckFile(group, f);
        }
    }
    
    private void CheckFile (String groupName, File f) {
    
            Workbook workbook;
            boolean updateFile = false;
        
            try {
                FileInputStream file = new FileInputStream(f);

		        workbook = WorkbookFactory.create(file);          
                Sheet sheet = workbook.getSheetAt(0);        
                
                String sheetGroupName = getCellStringValue(sheet, ExcelProcessor.GROUP_NAME_CELL);
                
                if (!groupName.equals(sheetGroupName)) {
                        System.out.println("!!! WARNING: Wrong group number/name");
                        System.out.println("Should be " + groupName + " but is " + sheetGroupName);
                        System.out.println("You wish to fix it? (y/n)");
                        String option = scanner.nextLine();
                        
                        if (option.equals("y")) {
                            updateFile = true;
                            getCellByCoordenates(sheet, ExcelProcessor.GROUP_NAME_CELL).setCellValue(groupName);
                        }
                }
                
                try {
                
                    int day = getCellIntValue(sheet,ExcelProcessor.DATE_DAY_CELL);
                    int month = getCellIntValue(sheet,ExcelProcessor.DATE_MONTH_CELL);
                    int year = getCellIntValue(sheet,ExcelProcessor.DATE_YEAR_CELL);
                    
                    if (day > 31 || day < 1 || month > 12 || month < 1 || year < 2010 || year > 2050)
                        throw new java.lang.NumberFormatException();
                    
                } catch (java.lang.NumberFormatException e) {
                
                    updateFile = true; 
                    
                    System.out.println("!!! ERROR: Date info is wrong.");
                    System.out.println("Hint: File name is " + f.getName());
                    
                    System.out.println("Set the day: ");
                    String day = scanner.nextLine();
                    getCellByCoordenates(sheet, ExcelProcessor.DATE_DAY_CELL).setCellValue(day);
                    
                    System.out.println("Set the month: ");
                    String month = scanner.nextLine();
                    getCellByCoordenates(sheet, ExcelProcessor.DATE_MONTH_CELL).setCellValue(month);
                    
                    System.out.println("Set the year: ");
                    String year = scanner.nextLine();
                    getCellByCoordenates(sheet, ExcelProcessor.DATE_YEAR_CELL).setCellValue(year);
                }
                
                file.close();
                
                if (updateFile) {
                    FileOutputStream outFile = new FileOutputStream(f);
                    workbook.write(outFile);
                    outFile.close();
                }
                
            } catch (Exception e) {
                e.printStackTrace();
                System.exit(0);
            }
    }
    
    private Cell getCellByCoordenates(Sheet sheet, String cellStr) {
        CellReference cellReference = new CellReference(cellStr);
        Row row = sheet.getRow(cellReference.getRow());
        return row.getCell(cellReference.getCol()); 
    }
    
    private String getCellStringValue(Sheet sheet, String cellStr) {
        CellReference cellReference = new CellReference(cellStr); 
        Row row = sheet.getRow(cellReference.getRow());
        Cell cell = row.getCell(cellReference.getCol()); 
        
        return getCellStringValue(cell);
    }
    
    private String getCellStringValue(Cell cell) {        
        switch (cell.getCellTypeEnum()) {
            case STRING: 
                return cell.getStringCellValue(); 
            case NUMERIC: 
                return Double.toString(cell.getNumericCellValue()); 
        }
        
        return "";
    }
    
    private int getCellIntValue(Sheet sheet, String cellStr) {
        CellReference cellReference = new CellReference(cellStr); 
        Row row = sheet.getRow(cellReference.getRow());
        Cell cell = row.getCell(cellReference.getCol()); 
        
        return getCellIntValue(cell);
    }
    
    private int getCellIntValue(Cell cell) {
        switch (cell.getCellTypeEnum()) {
            case STRING: 
                return Integer.parseInt(cell.getStringCellValue()); 
            case NUMERIC: 
                return (int) cell.getNumericCellValue(); 
        }
        return 0;
    }
}
