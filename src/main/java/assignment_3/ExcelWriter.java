package assignment_3;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelWriter {

    private static String[] columns = {"T", "|M|", "Do not need this", "chi", "cv"};
    private static List<Value> values = new ArrayList<>();


    // Initializing  data to insert into the excel file
    static {
        BufferedReader reader;
        try {
            reader = new BufferedReader(new FileReader(
                    "A350x50.txt"));
            String line = reader.readLine();
            while (line != null) {
                String[] arr = line.split(" ");
                double[] array = new double[arr.length];

                for (int i = 0; i < arr.length; i++) {
                    array[i] = Double.parseDouble(arr[i]);
                }

                Value value = new Value(array[0], array[1], 0.0, array[3], array[4]);

                values.add(value);

                // read next line
                line = reader.readLine();
            }
            reader.close();
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
        for (Value value : values) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(value.getT());
            row.createCell(1).setCellValue(value.getM());
            row.createCell(2).setCellValue(value.getMM_not_needed());
            row.createCell(3).setCellValue(value.getChi());
            row.createCell(4).setCellValue(value.getCv());
        }

        // Resize all columns to fit the content size
        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("result50.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
    }
}