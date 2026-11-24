package org.example;
import java.io.*;
import org.apache.poi.ss.usermodel.*;


public class Main {
    public static void main(String[] args) throws Exception  {
read();
    }
public static void read() throws IOException{
      try( FileInputStream inp = new FileInputStream("Inventory-Records-Sample-Data.xlsx")){
            Workbook workbook = WorkbookFactory.create(inp);
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();

            for (Row row : sheet){
                for (Cell cell : row) {
                    String value = formatter.formatCellValue(cell).replaceAll("\\s+", " ").trim();
                    System.out.printf("%-30s |", value);
                }
                System.out.println();
            }

        }

}}


