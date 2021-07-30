import java.util.LinkedHashMap;

public class DiemTotNghiep {
    public Double Toan;
    public Double Van;
    public Double TiengAnh;
    public Double Ly;
    public Double Hoa;
    public Double Sinh;
    public Double Su;
    public Double Dia;
    public Double GDCD;

    public LinkedHashMap<String, Double> DiemToHop = new LinkedHashMap<>();

    @Override
    public String toString() {
        return "DiemTotNghiep{" +
                "Toan=" + Toan +
                ", Van=" + Van +
                ", TiengAnh=" + TiengAnh +
                ", Ly=" + Ly +
                ", Hoa=" + Hoa +
                ", Sinh=" + Sinh +
                ", Su=" + Su +
                ", Dia=" + Dia +
                ", GDCD=" + GDCD +
                '}';
    }
}
