package org.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.excel.dao.SchoolClass;
import org.excel.dao.Student;
import org.excel.dao.Subject;
import org.excel.hardcoded.SchoolClassGenerator;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

public class Main {
    public static void main(String[] args) {
        List<SchoolClass> schoolClasses = SchoolClassGenerator.generateClasses(5, 20);
        List<String> headers = List.of("CLASS_ID", "CLASS_SUBJECT", "STUDENT_ID", "STUDENT_FIRST_NAME", "STUDENT_LAST_NAME", "STUDENT_AGE", "STUDENT_GRADES");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Class");
            CellStyle cellStyle = workbook.createCellStyle();
            Font font = workbook.createFont();
            createHeader(headers, sheet, cellStyle, font);
            populateWorkbook(sheet, schoolClasses, workbook.createCellStyle());
            autoSizeCells(sheet, headers);
            try (OutputStream os = new FileOutputStream("workbook.xlsx")) {
                workbook.write(os);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void autoSizeCells(Sheet sheet, List<String> headers) {
        for (int i = 0; i < headers.size(); i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private static void populateWorkbook(Sheet sheet, List<SchoolClass> classes, CellStyle cellStyle) {
        int ROW_CLASSES = 1;
        int AFTER_STUDENT_COUNTER = 1;
        cellStyle.setWrapText(true);
//        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        for (int idxClasses = 0; idxClasses < classes.size(); idxClasses++) {
            SchoolClass schoolClass = classes.get(idxClasses);
            Row classRow = sheet.createRow(ROW_CLASSES);
            Cell classIdCell = classRow.createCell(0, CellType.STRING);
            Cell classSubjectCell = classRow.createCell(1, CellType.STRING);
            classIdCell.setCellValue(schoolClass.getId().toString());
            classSubjectCell.setCellValue(schoolClass.getSubject());
//            classIdCell.setCellStyle(cellStyle);
//            classSubjectCell.setCellStyle(cellStyle);
            for (int idxStudent = 0; idxStudent < schoolClass.getStudents().size(); idxStudent++) {
                Student student = schoolClass.getStudents().get(idxStudent);
                if (idxStudent == 0) {
                    Cell studentIdCell = classRow.createCell(2, CellType.STRING);
                    Cell studentFirstNameCell = classRow.createCell(3, CellType.STRING);
                    Cell studentLastNameCell = classRow.createCell(4, CellType.STRING);
                    Cell studentAgeCell = classRow.createCell(5, CellType.STRING);
                    Cell studentGradesCell = classRow.createCell(6, CellType.STRING);
                    studentIdCell.setCellValue(student.getId().toString());
                    studentFirstNameCell.setCellValue(student.getFirstName());
                    studentLastNameCell.setCellValue(student.getLastName());
                    studentAgeCell.setCellValue(student.getAge());
                    studentGradesCell.setCellValue(toStringGrades(student.getGrades()));
                    studentGradesCell.setCellStyle(cellStyle);
//                    studentIdCell.setCellStyle(cellStyle);
//                    studentFirstNameCell.setCellStyle(cellStyle);
//                    studentLastNameCell.setCellStyle(cellStyle);
//                    studentAgeCell.setCellStyle(cellStyle);
                } else {
                    Row studentRow = sheet.createRow(ROW_CLASSES);
                    Cell studentIdCell = studentRow.createCell(2, CellType.STRING);
                    Cell studentFirstNameCell = studentRow.createCell(3, CellType.STRING);
                    Cell studentLastNameCell = studentRow.createCell(4, CellType.STRING);
                    Cell studentAgeCell = studentRow.createCell(5, CellType.STRING);
                    Cell studentGradesCell = studentRow.createCell(6, CellType.STRING);
                    studentIdCell.setCellValue(student.getId().toString());
                    studentFirstNameCell.setCellValue(student.getFirstName());
                    studentLastNameCell.setCellValue(student.getLastName());
                    studentAgeCell.setCellValue(student.getAge());
                    studentGradesCell.setCellValue(toStringGrades(student.getGrades()));
                    studentGradesCell.setCellStyle(cellStyle);
//                    studentIdCell.setCellStyle(cellStyle);
//                    studentFirstNameCell.setCellStyle(cellStyle);
//                    studentLastNameCell.setCellStyle(cellStyle);
//                    studentAgeCell.setCellStyle(cellStyle);
                }
                ROW_CLASSES++;
            }

        }
    }

    private static String toStringGrades(Map<Subject, List<Integer>> map) {
        var sb = new StringBuilder();
        map.forEach((key, value) -> {
            sb.append(key.toString()).append(" : ");
            value.forEach(v -> sb.append(v.toString()).append(", "));
            sb.append("\n");
        });
        return sb.toString();
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
        }
    }

}