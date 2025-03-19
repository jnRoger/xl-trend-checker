package org.example.xltrendchecker;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import static java.lang.Math.abs;

public class TrendChecker {
    public static void main(String[] args) throws IOException {

        //make dynamic?
        String filePath = "test.xlsx";

        Scanner scanner = new Scanner(System.in);
        System.out.println("Введите допустимое отклонение тренда.");

        //make float?
        int allowance = scanner.nextInt();

        int errorCount = 0;

        FileInputStream fileInputStream = new FileInputStream(filePath);

        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int rowCount = sheet.getLastRowNum();

        XSSFRow row = sheet.getRow(0);
        int cellCount = row.getLastCellNum();

        XSSFCellStyle errorCellStyle = workbook.createCellStyle();
        errorCellStyle.setFillForegroundColor(IndexedColors.PINK.getIndex());
        errorCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        for (int c = 1; c < cellCount; c++) {
            for (int r = 1; r < rowCount; r++) {

                row = sheet.getRow(r);
                XSSFCell cell = row.getCell(c);
                XSSFCell nextCell = sheet.getRow (r + 1).getCell(c);

                if ((abs(cell.getNumericCellValue() -
                        nextCell.getNumericCellValue())) > allowance) {
                    errorCount++;
                    nextCell.setCellStyle(errorCellStyle);
                }
            }
        }

        fileInputStream.close();
        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
        workbook.write(fileOutputStream);
        fileOutputStream.close();

        if (errorCount > 0)
            System.out.println("Тренд нестабилен, обнаружено " + errorCount
            + " скачков. Проверьте данные, отмеченные цветом.");
        else
            System.out.println("Тренд стабилен.");
    }
}
