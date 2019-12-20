package assignment_3;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelWriter {

    private static String[] columns = {"T", "|M|", "E", "chi", "cv"};
    private static List<Value> values10 = new ArrayList<>();
    private static List<Value> values20 = new ArrayList<>();
    private static List<Value> values30 = new ArrayList<>();
    private static List<Value> values40 = new ArrayList<>();
    private static List<Value> values50 = new ArrayList<>();

    // Initializing  data to insert into the excel file
    static {

    }

    public static void main(String[] args) {

        // Read the data from the txt files
        readDataFromTxt(values10, 10);
//        readDataFromTxt(values20, 20);
//        readDataFromTxt(values30, 30);
//        readDataFromTxt(values40, 40);
//        readDataFromTxt(values50, 50);

        // Write the data to an excel
        try {
            processData(values10, 10);
//            processData(values20,20);
//            processData(values30,30);
//            processData(values40,40);
//            processData(values50,50);

        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    private static void readDataFromTxt(List<Value> list, int number) {
        BufferedReader reader;
        try {
            reader = new BufferedReader(new FileReader(
                    "A3_" + number + "x" + number + ".txt"));
            String line = reader.readLine();
            while (line != null) {
                String[] arr = line.split(" ");
                double[] array = new double[arr.length];

                for (int i = 0; i < arr.length; i++) {
                    array[i] = Double.parseDouble(arr[i]);
                }

                Value value = new Value(array[0], array[1], array[2], array[3], array[4]);

                list.add(value);

                // read next line
                line = reader.readLine();
            }
            reader.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void processData(List<Value> valueList, int number) throws IOException {
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
        for (Value value : valueList) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(value.getT());
            row.createCell(1).setCellValue(value.getM());
            row.createCell(2).setCellValue(value.getE());
            row.createCell(3).setCellValue(value.getChi());
            row.createCell(4).setCellValue(value.getCv());
        }

        // Resize all columns to fit the content size
        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("Assignment3/Result_Total.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();

        System.out.println("Excel file with CONT = " + number + " created");
    }
}