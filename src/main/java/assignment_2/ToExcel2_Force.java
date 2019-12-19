package assignment_2;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ToExcel2_Force {

    private static String[] columns = {"Time", "F1", "F1.5", "F2", "F2.5", "F3"};
    private static List<Value2> values1 = new ArrayList<>();
    private static List<Value2> values15 = new ArrayList<>();
    private static List<Value2> values2 = new ArrayList<>();
    private static List<Value2> values25 = new ArrayList<>();
    private static List<Value2> values3 = new ArrayList<>();


    // Initializing  data to insert into the excel file
    static {

        // 1 reader
        BufferedReader reader1;
        try {
            reader1 = new BufferedReader(new FileReader(
                    "A2_1force.txt"));
            String line = reader1.readLine();
            while (line != null) {
                String[] arr = line.split(" ");
                double[] array = new double[arr.length];

                for (int i = 0; i < arr.length; i++) {
                    array[i] = Double.parseDouble(arr[i]);
                }

                Value2 value = new Value2(array[0], array[1]);

                values1.add(value);

                // read next line
                line = reader1.readLine();
            }
            reader1.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        // 1.5 reader
        BufferedReader reader15;
        try {
            reader15 = new BufferedReader(new FileReader(
                    "A2_15force.txt"));
            String line = reader15.readLine();
            while (line != null) {
                String[] arr = line.split(" ");
                double[] array = new double[arr.length];

                for (int i = 0; i < arr.length; i++) {
                    array[i] = Double.parseDouble(arr[i]);
                }

                Value2 value = new Value2(array[0], array[1]);

                values15.add(value);

                // read next line
                line = reader15.readLine();
            }
            reader15.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        // 2 reader
        BufferedReader reader2;
        try {
            reader2 = new BufferedReader(new FileReader(
                    "A2_2force.txt"));
            String line = reader2.readLine();
            while (line != null) {
                String[] arr = line.split(" ");
                double[] array = new double[arr.length];

                for (int i = 0; i < arr.length; i++) {
                    array[i] = Double.parseDouble(arr[i]);
                }

                Value2 value = new Value2(array[0], array[1]);

                values2.add(value);

                // read next line
                line = reader2.readLine();
            }
            reader2.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        // 2.5 reader
        BufferedReader reader25;
        try {
            reader25 = new BufferedReader(new FileReader(
                    "A2_25force.txt"));
            String line = reader25.readLine();
            while (line != null) {
                String[] arr = line.split(" ");
                double[] array = new double[arr.length];

                for (int i = 0; i < arr.length; i++) {
                    array[i] = Double.parseDouble(arr[i]);
                }

                Value2 value = new Value2(array[0], array[1]);

                values25.add(value);

                // read next line
                line = reader25.readLine();
            }
            reader25.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        // 3 reader
        BufferedReader reader3;
        try {
            reader3 = new BufferedReader(new FileReader(
                    "A2_3force.txt"));
            String line = reader3.readLine();
            while (line != null) {
                String[] arr = line.split(" ");
                double[] array = new double[arr.length];

                for (int i = 0; i < arr.length; i++) {
                    array[i] = Double.parseDouble(arr[i]);
                }

                Value2 value = new Value2(array[0], array[1]);

                values3.add(value);

                // read next line
                line = reader3.readLine();
            }
            reader3.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        // Create a Workbook
        Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

        /* CreationHelper helps us create instances of various things like DataFormat,
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet("Values");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Create cells
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
            cell.setCellStyle(workbook.createCellStyle());
        }

        // Create Other rows and cells with assignment_3.Value data
        int rowNum = 1;
        int time = 0;
        for (int i = 0; i < values1.size(); i++) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(time);
            time = time + 100;
            row.createCell(1).setCellValue(values1.get(i).getVelocity());
            row.createCell(2).setCellValue(values15.get(i).getVelocity());
            row.createCell(3).setCellValue(values2.get(i).getVelocity());
            row.createCell(4).setCellValue(values25.get(i).getVelocity());
            row.createCell(5).setCellValue(values3.get(i).getVelocity());
        }

        // Resize all columns to fit the content size
        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("Assignment2/Assignment2_Forces_Result.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();

        System.out.println("Excel table was build successfully :)");
    }
}
