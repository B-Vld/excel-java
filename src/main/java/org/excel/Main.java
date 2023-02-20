package org.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.excel.dao.SchoolClass;
import org.excel.hardcoded.SchoolClassGenerator;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
        List<SchoolClass> schoolClasses = SchoolClassGenerator.generateClasses(5, 20);
        List<String> headers = List.of("CLASS_ID", "CLASS_SUBJECT", "STUDENT_ID", "STUDENT_FIRST_NAME", "STUDENT_LAST_NAME", "STUDENT_AGE", "STUDENT_GRADES");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Class");
            CellStyle cellStyle = workbook.createCellStyle();
            Font font = workbook.createFont();
            createHeader(headers, sheet, cellStyle, font);
            populateWorkbook(sheet, schoolClasses);
            try (OutputStream os = new FileOutputStream("workbook.xlsx")) {
                workbook.write(os);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void populateWorkbook(Sheet sheet, List<SchoolClass> classes) {
        for(int idxClasses = 0; idxClasses < classes.size(); idxClasses++) {
            Row classRow = sheet.createRow(idxClasses + 1);

        }
    }

    private static void createHeader(List<String> headers, Sheet sheet, CellStyle cellStyle, Font font) {
        Row headerRow = sheet.createRow(0);
        font.setBold(true);
        font.setFontHeightInPoints((short) 15);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setBorderTop(BorderStyle.THICK);
        cellStyle.setBorderBottom(BorderStyle.THICK);
        cellStyle.setFont(font);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.createCell(i, CellType.STRING);
            cell.setCellValue(headers.get(i));
            cell.setCellStyle(cellStyle);
            sheet.autoSizeColumn(i);
        }
    }

}