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

public class ExcelProcessor {
    
    public static String GROUP_NAME_CELL = "E5";
    public static String TOWN_CELL = "E6";
    public static String PLACE_STD_CELL = "E7";
    public static String PLACE_NONSTD_CELL = "E8";
    public static String DATE_DAY_CELL = "E9";
    public static String DATE_MONTH_CELL = "F9";
    public static String DATE_YEAR_CELL = "G9";
    
    public static String GIRLS_PATOLOGIES_START_CELL = "C15";
    public static String WOMEN_PATOLOGIES_START_CELL = "C32";
    public static String MEN_PATOLOGIES_START_CELL = "Q32";
    public static String BOYS_PATOLOGIES_START_CELL = "Q15";
    
    public static String TALKS_START_CELL = "J5";
    public static String TALKS_START_NTIMES_CELL = "O5";
    public static String TALKS_START_DURATION_CELL = "P5";
    public static String TALKS_START_ADULTSORCHILDREN_CELL = "Q5";
    public static String TALKS_START_ATTENDEES_CELL = "R5";
    public static String TALKS_START_COMMENTS_CELL = "T5";
    public static String TALKS_END_CELL = "J10";
    
    public String folderAddr = "";
    public static String NAME_GENERATED_FILE = "GeneratedResultFile.xlsx";
    
    private ArrayList<DailyStatistic> statisticList;
    private ArrayList<File> statisticSheets;
    
    private ExcelIndexes eIndexes;
    
    private Workbook resultWorkbook;
    
    private Scanner scanner;
    
    public ExcelProcessor (Scanner scanner) {
        
        this.scanner = scanner;
        
        statisticList = new ArrayList<DailyStatistic>();
        statisticSheets = new ArrayList<File>();
        //scanner = new Scanner(System.in);
        
        System.out.println("********** Starting file parsing.");
        parseFiles();
        System.out.println("********** Successfully ended file parsing.");
        System.out.println("********** Starting data integrity check.");
        checkTownDataIntegrity();
        checkPlaceDataIntegrity();
        System.out.println("********** Successfully ended data integrity check.");
        System.out.println("********** Starting creating indexes.");
        createIndexes();
        System.out.println("********** Successfully ended creating indexes.");
        System.out.println("********** Starting writing result file.");
        writeResultFile();
        System.out.println("********** Successfully ended writing result file.");
        System.out.println("********** Starting saving result file.");
        saveResultFile();
        System.out.println("********** Successfully ended saving result file.");
        
        //scanner.close();
    }
    
    private void parseFiles() {
    
    System.out.println("Choose the target folder: ");
    folderAddr = scanner.nextLine();
        
	System.out.println("** Reading folder " + folderAddr  +  " **");

	File dir = new File(folderAddr);
    File[] filesOnDir = dir.listFiles();
	
	System.out.println("** " + filesOnDir.length  +  " files found **");
	
        for (File f : filesOnDir) {
            if (f.isFile() && f.getName().endsWith(".xlsx") && !f.getName().equals(NAME_GENERATED_FILE))
                statisticSheets.add(f);
        }
        
	int fileCounter = 1;

        for (File f : statisticSheets) {
            System.out.println("Parsing file " + fileCounter++  +  " \""+f+"\".");
            Workbook workbook;
        
            try {
                FileInputStream file = new FileInputStream(f);
                // workbook = new XSSFWorkbook(file);  
                workbook = WorkbookFactory.create(file);          
                Sheet sheet = workbook.getSheetAt(0);        
                
                DailyStatistic dailyStatistic = new DailyStatistic();
                dailyStatistic.groupName = getCellStringValue(sheet,GROUP_NAME_CELL);
                dailyStatistic.town = normalizeText(getCellStringValue(sheet,TOWN_CELL));
                dailyStatistic.place = getCellStringValue(sheet,PLACE_STD_CELL);
                if (dailyStatistic.place.equals("Otro"))            
                    dailyStatistic.place = normalizeText(getCellStringValue(sheet,PLACE_NONSTD_CELL));
                
                int day = getCellIntValue(sheet,DATE_DAY_CELL);
                int month = getCellIntValue(sheet,DATE_MONTH_CELL);
                int year = getCellIntValue(sheet,DATE_YEAR_CELL);
                dailyStatistic.date = LocalDate.of(year, month, day);
                
                PatologyBlock girlsPatologies = new PatologyBlock();
                parsePatologies(sheet,GIRLS_PATOLOGIES_START_CELL,girlsPatologies);
                girlsPatologies.computeTotal();
                PatologyBlock womenPatologies = new PatologyBlock();
                parsePatologies(sheet,WOMEN_PATOLOGIES_START_CELL,womenPatologies);
                womenPatologies.computeTotal();
                PatologyBlock boysPatologies = new PatologyBlock();
                parsePatologies(sheet,BOYS_PATOLOGIES_START_CELL,boysPatologies);
                boysPatologies.computeTotal();
                PatologyBlock menPatologies = new PatologyBlock();
                parsePatologies(sheet,MEN_PATOLOGIES_START_CELL,menPatologies);
                menPatologies.computeTotal();
                
                dailyStatistic.girlsPatologies = girlsPatologies;
                dailyStatistic.womenPatologies = womenPatologies;
                dailyStatistic.menPatologies = menPatologies;
                dailyStatistic.boysPatologies = boysPatologies;
                
                dailyStatistic.talks = new ArrayList<TalkBlock>();
                for (int row = getCellRowNumber(sheet, TALKS_START_CELL); row<getCellRowNumber(sheet, TALKS_END_CELL); row++) {
                    TalkBlock talk = new TalkBlock();
                    if (!getCellStringValue(getCellByCoordenates(sheet, row, getCellColNumber(sheet, TALKS_START_CELL))).equals("")) {
                        talk.subject = normalizeText(getCellStringValue(getCellByCoordenates(sheet, row, getCellColNumber(sheet, TALKS_START_CELL))));
                        talk.nTimes = getCellIntValue(getCellByCoordenates(sheet, row, getCellColNumber(sheet, TALKS_START_NTIMES_CELL)));
                        talk.duration = getCellIntValue(getCellByCoordenates(sheet, row, getCellColNumber(sheet, TALKS_START_DURATION_CELL)));
                        talk.childOrAdult = normalizeText(getCellStringValue(getCellByCoordenates(sheet, row, getCellColNumber(sheet, TALKS_START_ADULTSORCHILDREN_CELL))));
                        talk.attendees = getCellIntValue(getCellByCoordenates(sheet, row, getCellColNumber(sheet, TALKS_START_ATTENDEES_CELL)));
                        talk.comments = getCellStringValue(getCellByCoordenates(sheet, row, getCellColNumber(sheet, TALKS_START_COMMENTS_CELL)));
                        dailyStatistic.talks.add(talk);
                    }
                }                

                statisticList.add(dailyStatistic);
                
                file.close();
            } catch (Exception e) {
                e.printStackTrace();
                System.exit(0);
            }
        }
        
        Collections.sort(statisticList);
    }
    
    private void checkTownDataIntegrity() {    
        ArrayList<String> stringList;
        boolean add, correct = false;
        String answer;
        
        while (!correct) {
            stringList = new ArrayList<String>();
            for (DailyStatistic ds : statisticList) {
                add = true;
                for (String s : stringList) {
                    if (s.equals(ds.town))
                        add = false;
                }
                if (add)
                    stringList.add(ds.town);
            }
            
            System.out.println();
            System.out.println("Please check if all towns are correctly spelled.");
            
            for (int i = 0; i<stringList.size(); i++) {
                System.out.println("["+i+"] "+stringList.get(i));
            }
            
            System.out.print("Would you like to modify any element?: ");
            answer = scanner.nextLine();
            if (answer.equals("No") || answer.equals("no") || answer.equals("n") || answer.equals("N"))
                correct = true;
            else {
                try {
                    int index = Integer.parseInt(answer);
                    System.out.println("Write new name for "+stringList.get(index));
                    answer = scanner.nextLine();
                    
                    for (DailyStatistic ds : statisticList) {
                        if (ds.town.equals(stringList.get(index)))
                         ds.town = answer;
                    }
                } catch (NumberFormatException e) {
                    System.out.println("Command not recognised.");
                } catch (IndexOutOfBoundsException e) {
                    System.out.println("Unknown town.");                    
                }
            }
        }
    }
    
    private void checkPlaceDataIntegrity() {    
        ArrayList<String> stringList;
        boolean add, correct = false;
        String answer;
        
        while (!correct) {
            stringList = new ArrayList<String>();
            for (DailyStatistic ds : statisticList) {
                add = true;
                for (String s : stringList) {
                    if (s.equals(ds.place))
                        add = false;
                }
                if (add)
                    stringList.add(ds.place);
            }
            
            System.out.println();
            System.out.println("Please check if all places are correctly spelled.");
            
            for (int i = 0; i<stringList.size(); i++) {
                System.out.println("["+i+"] "+stringList.get(i));
            }
            
            System.out.print("Would you like to modify any element?: ");
            answer = scanner.nextLine();
            if (answer.equals("No") || answer.equals("no") || answer.equals("n") || answer.equals("N"))
                correct = true;
            else {
                try {
                    int index = Integer.parseInt(answer);
                    System.out.println("Write new name for "+stringList.get(index));
                    answer = scanner.nextLine();
                    
                    for (DailyStatistic ds : statisticList) {
                        if (ds.place.equals(stringList.get(index)))
                         ds.place = answer;
                    }
                } catch (NumberFormatException e) {
                    System.out.println("Command not recognised.");
                } catch (IndexOutOfBoundsException e) {
                    System.out.println("Unknown place.");                    
                }
            }
        }
    }
    
    private void createIndexes() {
    
        eIndexes = new ExcelIndexes();
        
        ArrayList<String> groups = new ArrayList<String>();
        ArrayList<String> towns = new ArrayList<String>();
        ArrayList<String> places = new ArrayList<String>();
        
        boolean addGroups;
        boolean addTowns;
        boolean addPlaces;
        
        for (DailyStatistic ds : statisticList) {
        
            addGroups = true;
            addTowns = true;
            addPlaces = true;
            
            for (String s : groups) {
                if (s.equals(ds.groupName))
                    addGroups = false;
            }
            
            for (String s : towns) {
                if (s.equals(ds.town))
                    addTowns = false;
            }
            
            for (String s : places) {
                if (s.equals(ds.place))
                    addPlaces = false;
            }
            
            if (addGroups)
                groups.add(ds.groupName);
            
            if (addTowns)
                towns.add(ds.town);
            
            if (addPlaces)
                places.add(ds.place);
        }
        
        Collections.sort(groups);
        Collections.sort(towns);
        Collections.sort(places);
        
        eIndexes.groups = groups;
        eIndexes.towns = towns;
        eIndexes.places = places;
    }
    
    private void writeResultFile() {
        resultWorkbook = new XSSFWorkbook();
        System.out.println("Creating sheet Poblaciones");
        Sheet poblaciones = resultWorkbook.createSheet("Poblaciones");
        writePoblacionesSheet(poblaciones);
        System.out.println("Creating sheet Patologías");
        Sheet patologias = resultWorkbook.createSheet("Patologias");
        writePatologiasSheet(patologias);
        System.out.println("Creating sheet Charlas");
        Sheet charlas = resultWorkbook.createSheet("Charlas");
        writeCharlasSheet(charlas);
    }
    
    private void writePoblacionesSheet(Sheet sheet) {
    
        int currentRow = 0;
        Cell c;
        Row row;
        
        for (String group : eIndexes.groups) {   
            int totalGroupGirls = 0, totalGroupBoys = 0, totalGroupWomen = 0, totalGroupMen = 0;
            row = sheet.createRow(currentRow);
            for (int i = 0; i < 8; i++)
                row.createCell(i);
            sheet.addMergedRegion(new CellRangeAddress(currentRow, currentRow, 0, 7));
            getCellByCoordenates(sheet, currentRow, 0).setCellValue(group.toUpperCase());
            
            row = sheet.createRow(++currentRow);
            c = row.createCell(0);
            c.setCellValue("Fecha");
            c = row.createCell(1);
            c.setCellValue("Municipio");
            c = row.createCell(2);
            c.setCellValue("Lugar");
            c = row.createCell(3);
            c.setCellValue("Niñas");
            c = row.createCell(4);
            c.setCellValue("Niños");
            c = row.createCell(5);
            c.setCellValue("Mujeres");
            c = row.createCell(6);
            c.setCellValue("Hombres");
            c = row.createCell(7);
            c.setCellValue("Total");
            
            for (DailyStatistic ds : statisticList) {
                if (!ds.groupName.equals(group))
                    continue;
                    
                row = sheet.createRow(++currentRow);
                c = row.createCell(0);
                c.setCellValue(ds.date.toString());
                c = row.createCell(1);
                c.setCellValue(ds.town);
                c = row.createCell(2);
                c.setCellValue(ds.place);
                c = row.createCell(3);
                c.setCellValue(ds.girlsPatologies.total);
                totalGroupGirls += ds.girlsPatologies.total;
                c = row.createCell(4);
                c.setCellValue(ds.boysPatologies.total);
                totalGroupBoys += ds.boysPatologies.total;
                c = row.createCell(5);
                c.setCellValue(ds.womenPatologies.total);
                totalGroupWomen += ds.womenPatologies.total;
                c = row.createCell(6);
                c.setCellValue(ds.menPatologies.total);
                totalGroupMen += ds.menPatologies.total;
                c = row.createCell(7);
                c.setCellValue(ds.girlsPatologies.total + ds.boysPatologies.total + ds.womenPatologies.total + ds.menPatologies.total);
            }
            
            row = sheet.createRow(++currentRow);
            c = row.createCell(2);
            c.setCellValue("Total");
            c = row.createCell(3);
            c.setCellValue(totalGroupGirls);
            c = row.createCell(4);
            c.setCellValue(totalGroupBoys);
            c = row.createCell(5);
            c.setCellValue(totalGroupWomen);
            c = row.createCell(6);
            c.setCellValue(totalGroupMen);
            c = row.createCell(7);
            int totalGroupAll = totalGroupGirls + totalGroupBoys + totalGroupWomen + totalGroupMen;
            c.setCellValue(totalGroupAll);
            
            row = sheet.createRow(++currentRow);
            c = row.createCell(3);
            c.setCellValue((totalGroupGirls * 100 / totalGroupAll) + " %");
            c = row.createCell(4);
            c.setCellValue((totalGroupBoys * 100 / totalGroupAll) + " %");
            c = row.createCell(5);
            c.setCellValue((totalGroupWomen * 100 / totalGroupAll) + " %");
            c = row.createCell(6);
            c.setCellValue((totalGroupMen * 100 / totalGroupAll) + " %");
            
            currentRow += 5;
        }
    }
    
    private void writePatologiasSheet(Sheet sheet) {
        int startRow, currentRow = 0, currentCol;
        Cell c;
        Row row;
        
        for (String group : eIndexes.groups) {  
        
            int[] totalGroupPatologies = new int [ExcelStrUtil.patogolyIds.length];
            for (int i = 0; i <totalGroupPatologies.length; i++) 
                totalGroupPatologies[i] = 0;
                    
            startRow = currentRow;
            
            row = sheet.createRow(currentRow++);
            row = sheet.createRow(currentRow++);
            c = row.createCell(0);
            c.setCellValue("ID");
            c = row.createCell(1);
            c.setCellValue("Patología");
            
            for (int i = 0; i < ExcelStrUtil.patogolyIds.length; i++) {
                row = sheet.createRow(currentRow++);
                c = row.createCell(0);
                c.setCellValue(ExcelStrUtil.patogolyIds[i]);
                c = row.createCell(1);
                c.setCellValue(ExcelStrUtil.patogolyNames[i]);
            }
            
            row = sheet.createRow(currentRow++);
            c = row.createCell(1);
            c.setCellValue("Total");
            
            currentCol = 2;
            
            for (DailyStatistic ds : statisticList) {
                if (!ds.groupName.equals(group))
                    continue;
                    
                int[] totalDayPatologies = new int [ds.girlsPatologies.retrievePatologies().size()];
                for (int i = 0; i <totalDayPatologies.length; i++) 
                    totalDayPatologies[i] = 0;
                    
                currentRow = startRow;
                
                row = sheet.getRow(currentRow++);
                c = row.createCell(currentCol);
                c.setCellValue(ds.date.toString());
                sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, currentCol, currentCol + 4));
                
                row = sheet.getRow(currentRow++);
                c = row.createCell(currentCol);
                c.setCellValue("Niñas");                
                
                ArrayList<Integer> girlsPatologies = ds.girlsPatologies.retrievePatologies();
                for (int i = 0; i < girlsPatologies.size(); i++) {
                    row = sheet.getRow(currentRow++);
                    c = row.createCell(currentCol);
                    c.setCellValue(girlsPatologies.get(i).intValue());
                    totalDayPatologies[i] += girlsPatologies.get(i).intValue();
                }
                
                row = sheet.getRow(currentRow++);
                c = row.createCell(currentCol);
                c.setCellValue(ds.girlsPatologies.total);
                
                currentRow = startRow;
                currentCol++; 
                
                row = sheet.getRow(currentRow++);
                row = sheet.getRow(currentRow++);
                c = row.createCell(currentCol);
                c.setCellValue("Niños");
                
                ArrayList<Integer> boysPatologies = ds.boysPatologies.retrievePatologies();
                for (int i = 0; i < boysPatologies.size(); i++) {
                    row = sheet.getRow(currentRow++);
                    c = row.createCell(currentCol);
                    c.setCellValue(boysPatologies.get(i).intValue());
                    totalDayPatologies[i] += boysPatologies.get(i).intValue();
                }
                
                row = sheet.getRow(currentRow++);
                c = row.createCell(currentCol);
                c.setCellValue(ds.boysPatologies.total);
                
                currentRow = startRow;
                currentCol++;
                
                row = sheet.getRow(currentRow++);
                row = sheet.getRow(currentRow++);
                c = row.createCell(currentCol);
                c.setCellValue("Mujeres");
                
                ArrayList<Integer> womenPatologies = ds.womenPatologies.retrievePatologies();
                for (int i = 0; i < womenPatologies.size(); i++) {
                    row = sheet.getRow(currentRow++);
                    c = row.createCell(currentCol);
                    c.setCellValue(womenPatologies.get(i).intValue());
                    totalDayPatologies[i] += womenPatologies.get(i).intValue();
                }
                
                row = sheet.getRow(currentRow++);
                c = row.createCell(currentCol);
                c.setCellValue(ds.womenPatologies.total);
                
                currentRow = startRow;
                currentCol++;
                
                row = sheet.getRow(currentRow++);
                row = sheet.getRow(currentRow++);
                c = row.createCell(currentCol);
                c.setCellValue("Hombres");
                
                ArrayList<Integer> menPatologies = ds.menPatologies.retrievePatologies();
                for (int i = 0; i < menPatologies.size(); i++) {
                    row = sheet.getRow(currentRow++);
                    c = row.createCell(currentCol);
                    c.setCellValue(menPatologies.get(i).intValue());
                    totalDayPatologies[i] += menPatologies.get(i).intValue();
                } 
                
                row = sheet.getRow(currentRow++);
                c = row.createCell(currentCol);
                c.setCellValue(ds.menPatologies.total);
                
                currentRow = startRow;
                currentCol++;
                
                row = sheet.getRow(currentRow++);
                row = sheet.getRow(currentRow++);
                c = row.createCell(currentCol);
                c.setCellValue("Total");
                
                int totalDayAll = 0;
                
                for (int i = 0; i < totalDayPatologies.length; i++) {
                    row = sheet.getRow(currentRow++);
                    c = row.createCell(currentCol);
                    c.setCellValue(totalDayPatologies[i]);
                    totalGroupPatologies[i] += totalDayPatologies[i];
                    totalDayAll += totalDayPatologies[i];
                }    
                
                row = sheet.getRow(currentRow++);
                c = row.createCell(currentCol);
                c.setCellValue(totalDayAll);
                
                currentCol++;                
            }
            
            currentRow = startRow;
            currentCol++;
            
            row = sheet.getRow(currentRow++);
            row = sheet.getRow(currentRow++);
            c = row.createCell(currentCol);
            c.setCellValue("TOTAL");
            
            int totalGroupAll = 0;
            
            for (int i = 0; i < totalGroupPatologies.length; i++) {
                row = sheet.getRow(currentRow++);
                c = row.createCell(currentCol);
                c.setCellValue(totalGroupPatologies[i]);
                totalGroupAll += totalGroupPatologies[i];
            }    
            
            row = sheet.getRow(currentRow++);
            c = row.createCell(currentCol);
            c.setCellValue(totalGroupAll);
            
            currentCol++;   
            
            currentRow += 5;
        }
    }
    
    private void writeCharlasSheet(Sheet sheet) {
        int startRow, currentRow = 0, currentCol;
        Cell c;
        Row row;
        
        for (String group : eIndexes.groups) {
                
            int totalGroupTalks = 0, totalGroupAttendees = 0;
            
            row = sheet.createRow(currentRow);
            c = row.createCell(0);
            c.setCellValue("Fecha");
            c = row.createCell(1);
            c.setCellValue("Municipio");
            c = row.createCell(2);
            c.setCellValue("Lugar");
            c = row.createCell(3);
            c.setCellValue("Tema");
            c = row.createCell(4);
            c.setCellValue("N Veces");
            c = row.createCell(5);
            c.setCellValue("Duración");
            c = row.createCell(6);
            c.setCellValue("Adultos o niños");
            c = row.createCell(7);
            c.setCellValue("Asistentes");
            c = row.createCell(8);
            c.setCellValue("Comentarios");
            
            currentRow++;
                        
            for (DailyStatistic ds : statisticList) {  
                if (!ds.groupName.equals(group))
                    continue;
                    
                int totalDayTalks = 0, totalDayAttendees = 0;
                for (TalkBlock talk : ds.talks) {
                    row = sheet.createRow(currentRow);
                    
                    c = row.createCell(0);
                    c.setCellValue(ds.date.toString());
                    
                    c = row.createCell(1);
                    c.setCellValue(ds.town);
                    
                    c = row.createCell(2);
                    c.setCellValue(ds.place);
                    
                    c = row.createCell(3);
                    c.setCellValue(talk.subject);
                    
                    c = row.createCell(4);
                    c.setCellValue(talk.nTimes);
                    totalDayTalks += talk.nTimes;
                    
                    c = row.createCell(5);
                    c.setCellValue(talk.duration);
                    
                    c = row.createCell(6);
                    c.setCellValue(talk.childOrAdult);
                    
                    c = row.createCell(7);
                    c.setCellValue(talk.attendees);
                    totalDayAttendees += talk.attendees;
                    
                    c = row.createCell(8);
                    c.setCellValue(talk.comments);
                    
                    currentRow++;
                }
                
                row = sheet.createRow(currentRow);
                c = row.createCell(3);
                c.setCellValue("Total:");
                c = row.createCell(4);
                c.setCellValue(totalDayTalks);
                c = row.createCell(6);
                c.setCellValue("Total:");
                c = row.createCell(7);
                c.setCellValue(totalDayAttendees);
                
                currentRow++;       
                
                totalGroupTalks += totalDayTalks;
                totalGroupAttendees += totalDayAttendees;         
            }
            
            currentRow++;  
            
            row = sheet.createRow(currentRow);
            c = row.createCell(3);
            c.setCellValue("TOTAL:");
            c = row.createCell(4);
            c.setCellValue(totalGroupTalks);
            c = row.createCell(6);
            c.setCellValue("TOTAL:");
            c = row.createCell(7);
            c.setCellValue(totalGroupAttendees);
            
            currentRow++;  
            
            currentRow += 5;
        }
    }
    
    private void saveResultFile() {
        try {
            FileOutputStream fileout = new FileOutputStream(folderAddr+NAME_GENERATED_FILE);
            resultWorkbook.write(fileout);
            fileout.close();
            resultWorkbook.close();
        } catch (Exception e) {
            e.printStackTrace();
            System.exit(0);
        }
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
    
    private void parsePatologies(Sheet sheet, String cellStr, PatologyBlock patologyBlock) {
        int startRow, currentRow;
        int startCol, currentCol;
        CellReference cellReference;
        Cell currentCell;
        
        cellReference = new CellReference(cellStr);
        startRow = cellReference.getRow();
        startCol = cellReference.getCol();
        
        currentRow = startRow;
        currentCol = startCol;
        
        boolean exitLoop = false;
        while (!exitLoop) {
            currentCell = getCellByCoordenates(sheet, currentRow, currentCol);
            if (cellIsEmpty(currentCell)) {
                currentRow++;
                currentCol = startCol;
                currentCell = getCellByCoordenates(sheet, currentRow, currentCol);
                if (cellIsEmpty(currentCell))
                    break;
                else
                    continue;
            }
            patologyBlock.setPatology (getCellStringValue(currentCell), getCellIntValue(getCellByCoordenates(sheet, currentRow, currentCol+1)));
            currentCol = currentCol + 2;
        }        
    }
    
    private String normalizeText(String input) {
        String upper, lower;
        
        if (input.length() <= 0) return input;

        upper = input.substring(0,1).toUpperCase();
        lower = input.substring(1,input.length()).toLowerCase();
        
        return upper + lower;
    }
    
    private Cell getCellByCoordenates(Sheet sheet, int rowR, int colR) {
        CellReference cellReference = new CellReference(rowR, colR); 
        Row row = sheet.getRow(cellReference.getRow());
        return row.getCell(cellReference.getCol()); 
    }   
    
    private int getCellColNumber(Sheet sheet, String cellStr) {
        CellReference cellReference = new CellReference(cellStr);
        return cellReference.getCol(); 
    }
    
    private int getCellRowNumber(Sheet sheet, String cellStr) {
        CellReference cellReference = new CellReference(cellStr);
        return cellReference.getRow(); 
    }
    
    
    private boolean cellIsEmpty(Cell cell) {
        if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
            return true;
        }

        if (cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getStringCellValue().isEmpty()) {
            return true;
        }

        return false;
    }
}
