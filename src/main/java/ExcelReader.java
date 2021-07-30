
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.Serializable;
import java.util.*;

public class ExcelReader {

    public LinkedHashMap<String, Student> readExcel(String filename) throws IOException {
        LinkedHashMap<String, Student> list = new LinkedHashMap<>();
        FileInputStream fis = new FileInputStream(filename);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);
        FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
        int stt = 0;
        for (Row row : sheet) {
            if (stt++ == 0) continue;
            Student student = new Student();
            int col = 0;
            student.sobaodanh = getValue(formulaEvaluator, row.getCell(col++)).toString();
            list.put(student.sobaodanh, student);
        }
        return list;
    }

    int getIntegerValue(FormulaEvaluator formulaEvaluator, Cell cell) {
        if (formulaEvaluator.evaluateInCell(cell) == null)
            return 0;
        else
            switch (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum()) {
                case NUMERIC:
                    return (int) Math.round(cell.getNumericCellValue());
                case STRING:
                    if (cell.getStringCellValue().isEmpty())
                        return 0;
                    else
                        return Integer.parseInt(cell.getStringCellValue());
            }
        return 0;
    }


    Double getDoubleValue(FormulaEvaluator formulaEvaluator, Cell cell) {
        if (formulaEvaluator.evaluateInCell(cell) == null)
            return 0.0;
        else
            switch (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum()) {
                case NUMERIC:
                    return cell.getNumericCellValue();
                case STRING:
                    return Double.parseDouble(cell.getStringCellValue());
                case BLANK:
                    return 0.0;
            }
        return 0.0;
    }

    Double getDoubleValue1(FormulaEvaluator formulaEvaluator, Cell cell) {
        if (formulaEvaluator.evaluateInCell(cell) == null)
            return null;
        else
            switch (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum()) {
                case NUMERIC:
                    return cell.getNumericCellValue();
                case STRING:
                    if (cell.getStringCellValue().trim().isEmpty())
                        return 0.0;
                    else
                        return Double.parseDouble(cell.getStringCellValue());
                case BLANK:
                    return 0.0;
            }
        return null;
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

