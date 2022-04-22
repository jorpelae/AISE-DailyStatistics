package com.belwar.aiseexcelprocessor;
    
import java.util.Scanner;

public class MainClass {
    public static void main (String[] args) {
        Scanner scanner;
        scanner = new Scanner(System.in);
        System.out.println("Choose the program behaviour (check, analyze): ");
        String option = scanner.nextLine();
        
        if (option.equals("analyze") || option.equals("a"))
            new ExcelProcessor(scanner);
        else if (option.equals("check") || option.equals("c"))
            new ExcelChecker(scanner);
        else
            System.out.println("Unknown option: " + option);            
            
        scanner.close();
    }
}
