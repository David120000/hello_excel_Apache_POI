package dev.rd;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    
    public static void main(String[] args) {
        
        var data = readInputs();
        var workbook = createExcelWorkbook(data);
        writeToFile(workbook);

    }

    public static List<List<Object>> readInputs() {

        List<List<Object>> lines = new ArrayList<>();

        Scanner sc = new Scanner(System.in);
        boolean listeningInputs = true;

        do {
            System.out.println("Do you want to add a new message? Y/N");
            System.out.print("> ");
            String newLine = sc.nextLine();

            if(newLine.equals("Y") || newLine.equals("y")) {

                List<Object> line = new ArrayList<>();
                line.add(new Date());

                System.out.println("Your name:");
                System.out.print("> ");
                String name = sc.nextLine();
                line.add(name);
        
                System.out.println("Your message:");
                System.out.print("> ");
                String message = sc.nextLine();
                line.add(message);

                lines.add(line);
            }
            else if(newLine.equals("N") || newLine.equals("n")) {
                listeningInputs = false;
            }

        } while(listeningInputs);

        sc.close();

        return lines;
    }

    public static XSSFWorkbook createExcelWorkbook(List<List<Object>> data) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("messages");
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 4000);

        createHeaderRow(workbook, sheet);

        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setWrapText(true);

        int rowNumber = 1;
        for(var dataRow : data) {

            XSSFRow tableRow = sheet.createRow(rowNumber);
            rowNumber++;

            for(int i = 0; i < dataRow.size(); i++) {
                XSSFCell cell = tableRow.createCell(i);
                cell.setCellValue(dataRow.get(i).toString());
                cell.setCellStyle(cellStyle);
            }
        }

        return workbook;
    }

    private static void createHeaderRow(XSSFWorkbook workbook, XSSFSheet sheet) {
        
        XSSFRow headerRow = sheet.createRow(0);

        XSSFCellStyle headerStyle = workbook.createCellStyle();
        var color = new XSSFColor();
        color.setTheme(1);
        headerStyle.setFillBackgroundColor(color);

        XSSFFont font = workbook.createFont();
        font.setFontName("Calibri");
        font.setFontHeightInPoints(Short.valueOf("16"));
        font.setBold(true);
        headerStyle.setFont(font);

        XSSFCell cell0 = headerRow.createCell(0);
        cell0.setCellValue("Timestamp");
        cell0.setCellStyle(headerStyle);

        XSSFCell cell1 = headerRow.createCell(1);
        cell1.setCellValue("Name");
        cell1.setCellStyle(headerStyle);

        XSSFCell cell2 = headerRow.createCell(2);
        cell2.setCellValue("Message");
        cell2.setCellStyle(headerStyle);
    }

    public static void writeToFile(XSSFWorkbook workbook) {

        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String fileLocation = path.substring(0, path.length() - 1) + "temp.xlsx";

        try {
            FileOutputStream outputStream = new FileOutputStream(fileLocation);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        catch (IOException e) {
            e.printStackTrace();
        }
    }
}