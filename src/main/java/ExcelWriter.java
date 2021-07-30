
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.Serializable;
import java.util.LinkedHashMap;

public class ExcelWriter {

    public void writeExcel(LinkedHashMap<String, Student> students, String in, String out) throws IOException {
        FileInputStream fis = new FileInputStream(in);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);
        FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        CellStyle headerCellStyle = wb.createCellStyle();
        headerCellStyle.setFont(headerFont);
        Row headerRow = sheet.getRow(0);
        int stt = 1;
        int loai = 0;

        for (Row row : sheet) {
            if (stt++ <= 1) continue;  //Skip First Row
            String sbd = getValue(formulaEvaluator, row.getCell(0)).toString();
            Student student = students.get(sbd);
            int col = 1;
            if (student.diemtotnghiep.Toan != null)
                row.createCell(col++).setCellValue(String.format("%.2f", student.diemtotnghiep.Toan));
            else
                col++;
            if (student.diemtotnghiep.Van != null)
                row.createCell(col++).setCellValue(String.format("%.2f", student.diemtotnghiep.Van));
            else
                col++;
            if (student.diemtotnghiep.TiengAnh != null)
                row.createCell(col++).setCellValue(String.format("%.2f", student.diemtotnghiep.TiengAnh));
            else
                col++;
            if (student.diemtotnghiep.Ly != null)
                row.createCell(col++).setCellValue(String.format("%.2f", student.diemtotnghiep.Ly));
            else
                col++;
            if (student.diemtotnghiep.Hoa != null)
                row.createCell(col++).setCellValue(String.format("%.2f", student.diemtotnghiep.Hoa));
            else
                col++;
            if (student.diemtotnghiep.Sinh != null)
                row.createCell(col++).setCellValue(String.format("%.2f", student.diemtotnghiep.Sinh));
            else
                col++;
            if (student.diemtotnghiep.Su != null)
                row.createCell(col++).setCellValue(String.format("%.2f", student.diemtotnghiep.Su));
            else
                col++;
            if (student.diemtotnghiep.Dia != null)
                row.createCell(col++).setCellValue(String.format("%.2f", student.diemtotnghiep.Dia));
            else
                col++;

            if (student.diemtotnghiep.GDCD != null)
                row.createCell(col++).setCellValue(String.format("%.2f", student.diemtotnghiep.GDCD));
            else
                col++;
        }

        FileOutputStream fileOut = new FileOutputStream(out);
        wb.write(fileOut);
        fileOut.close();
        wb.close();
        System.out.println("Xong");
    }

    Serializable getValue(FormulaEvaluator formulaEvaluator, Cell cell) {
        if (formulaEvaluator.evaluateInCell(cell) == null)
            return null;
        else
            switch (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum()) {
                case NUMERIC:
                    return cell.getNumericCellValue();
                case STRING:
                    return cell.getStringCellValue();
            }
        return null;
    }
}
