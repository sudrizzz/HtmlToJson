import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.util.LinkedHashMap;

public class Main {

    public static void main(String[] args) {
        StringBuffer sb = new StringBuffer();
        String filepath = "layout6.html";
        String encoding = "UTF-8";

        try {
            // 读取文件将内容转成字符串
            FileInputStream fin = new FileInputStream(filepath);
            InputStreamReader reader = new InputStreamReader(fin, encoding);
            BufferedReader buffReader = new BufferedReader(reader);
            String strTmp;
            while ((strTmp = buffReader.readLine()) != null) {
                sb.append(strTmp + "\r\n");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        CellAnalysis cellAnalysis = new CellAnalysis();
        ExcelSecurity excelSecurity = new ExcelSecurity();
        LinkedHashMap map = cellAnalysis.getTableInfo(sb);
        System.out.println(map.get("datajson"));
        System.out.println(map.get("pluginjson"));
        System.out.println(map.get("script"));
    }
}
