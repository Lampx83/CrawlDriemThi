
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelWriter {

    public void writeExcel(String in, String out) throws IOException {

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

        String[] columns = {
                "A00", "A01", "B00", "C03", "C04", "D01", "D07", "D09", "D10",
                "A01x2", "D01x2", "D07x2", "D09x2", "D10x2",
                "ĐT", "Điểm XT", "ĐT", "Điểm XT", "ĐT", "Điểm XT", "ĐT", "Điểm XT", "ĐT", "Điểm XT",
                "TT Mã ngành", "TT ĐT", "TT NV", "TT Điểm", "TT Tên ngành", "Loại"
        };
    }
}
