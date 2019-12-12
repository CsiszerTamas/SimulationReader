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

public class ToExcel2 {

    private static String[] columns = {"Time", "AV100", "AV300", "AV500", "AV700"};
    private static List<Value2> values100 = new ArrayList<>();
    private static List<Value2> values300 = new ArrayList<>();
    private static List<Value2> values500 = new ArrayList<>();
    private static List<Value2> values700 = new ArrayList<>();


    // Initializing  data to insert into the excel file
    static {
        // first reader
        BufferedReader reader;
        try {
            reader = new BufferedReader(new FileReader(
                    "100pinning.txt"));
            String line = reader.readLine();
            while (line != null) {
                String[] arr = line.split(" ");
                double[] array = new double[arr.length];

                for (int i = 0; i < arr.length; i++) {
                    array[i] = Double.parseDouble(arr[i]);
                }

                Value2 value = new Value2(array[0], array[1]);

                values100.add(value);

                // read next line
                line = reader.readLine();
            }
            reader.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        //second reader
        BufferedReader reader2;
        try {
            reader2 = new BufferedReader(new FileReader(
                    "300pinning.txt"));
            String line = reader2.readLine();
            while (line != null) {
                String[] arr = line.split(" ");
                double[] array = new double[arr.length];

                for (int i = 0; i < arr.length; i++) {
                    array[i] = Double.parseDouble(arr[i]);
                }

                Value2 value = new Value2(array[0], array[1]);

                values300.add(value);

                // read next line
                line = reader2.readLine();
            }
            reader2.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        //third reader
        BufferedReader reader3;
        try {
            reader3 = new BufferedReader(new FileReader(
                    "500pinning.txt"));
            String line = reader3.readLine();
            while (line != null) {
                String[] arr = line.split(" ");
                double[] array = new double[arr.length];

                for (int i = 0; i < arr.length; i++) {
                    array[i] = Double.parseDouble(arr[i]);
                }

                Value2 value = new Value2(array[0], array[1]);

                values500.add(value);

                // read next line
                line = reader3.readLine();
            }
            reader3.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        //fourth reader
        BufferedReader reader4;
        try {
            reader4 = new BufferedReader(new FileReader(
                    "700pinning.txt"));
            String line = reader4.readLine();
            while (line != null) {
                String[] arr = line.split(" ");
                double[] array = new double[arr.length];

                for (int i = 0; i < arr.length; i++) {
                    array[i] = Double.parseDouble(arr[i]);
                }

                Value2 value = new Value2(array[0], array[1]);

                values700.add(value);

                // read next line
                line = reader4.readLine();
            }
            reader4.close();
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

        // Create Other rows and cells with Value data
        int rowNum = 1;
        int time = 0;
        for (int i = 0; i < values100.size(); i++) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(time);
            time = time + 100;
            row.createCell(1).setCellValue(values100.get(i).getVelocity());
            row.createCell(2).setCellValue(values300.get(i).getVelocity());
            row.createCell(3).setCellValue(values500.get(i).getVelocity());
            row.createCell(4).setCellValue(values700.get(i).getVelocity());
        }

        // Resize all columns to fit the content size
        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("Assignment2_result.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
    }
}