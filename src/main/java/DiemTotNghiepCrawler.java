
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

public class DiemTotNghiepCrawler {
    private HashSet<String> links;

    public DiemTotNghiepCrawler() {
        links = new HashSet<String>();
    }

    public DiemTotNghiep getPageLinks(int i, String URL) {
        DiemTotNghiep diemTotNghiep = new DiemTotNghiep();
        if (!links.contains(URL)) {
            try {
                if (links.add(URL)) {
                    // System.out.println(URL);
                }
                //2. Fetch the HTML code
                Document document = Jsoup.connect(URL).get();
                //3. Parse the HTML to extract links to other URLs

                Elements table = document.getElementsByClass("search-result").get(0).getElementsByClass("d-flex justify-content-between search-result-line py-3 px-3");

                if (table != null) {
                    for (Element element : table) {
                        if (element.toString().contains("Toán")) {
                            diemTotNghiep.Toan = Double.parseDouble(element.getElementsByClass("font-weight-bold").text());
                        } else if (element.toString().contains("Lí")) {
                            diemTotNghiep.Ly = Double.parseDouble(element.getElementsByClass("font-weight-bold").text());
                        } else if (element.toString().contains("Hóa")) {
                            diemTotNghiep.Hoa = Double.parseDouble(element.getElementsByClass("font-weight-bold").text());
                        } else if (element.toString().contains("Sinh")) {
                            diemTotNghiep.Sinh = Double.parseDouble(element.getElementsByClass("font-weight-bold").text());
                        } else if (element.toString().contains("Văn")) {
                            diemTotNghiep.Van = Double.parseDouble(element.getElementsByClass("font-weight-bold").text());
                        } else if (element.toString().contains("Sử")) {
                            diemTotNghiep.Su = Double.parseDouble(element.getElementsByClass("font-weight-bold").text());
                        } else if (element.toString().contains("Địa")) {
                            diemTotNghiep.Dia = Double.parseDouble(element.getElementsByClass("font-weight-bold").text());
                        } else if (element.toString().contains("N1")) {
                            diemTotNghiep.TiengAnh = Double.parseDouble(element.getElementsByClass("font-weight-bold").text());
                        }
                    }
                    System.err.println(i + " " + URL + "': OK");
                } else {
                    System.err.println(i + " " + URL + "': Không có thí sinh trên");
                }
            } catch (Exception e) {
                diemTotNghiep.Toan = -1.0;  //Error
            }
        }
        return diemTotNghiep;
    }

    public static void main(String[] args) throws InterruptedException {
        int from = 10000;
        int to = 619999;
        ArrayBlockingQueue<Runnable> queue = new ArrayBlockingQueue<>(to - from);
        ThreadPoolExecutor pool = new ThreadPoolExecutor(30, 60, 1, TimeUnit.HOURS, queue);
        for (int i = 10000; i <= to; i++) {
            int finalI = i;
            pool.execute(() -> {
                DiemTotNghiep diemtotnghiep = new DiemTotNghiepCrawler().getPageLinks(finalI, "https://vietnamnet.vn/vn/giao-duc/tra-cuu-diem-thi-thpt/?y=2021&sbd=" + finalI);
                if (diemtotnghiep != null) {

                }
            });
        }

        pool.shutdown();
        while (!pool.awaitTermination(5, TimeUnit.SECONDS)) {
        }
    }
}