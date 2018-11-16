import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.util.LinkedHashMap;

public class Main {

    public static void main(String[] args) {
        StringBuffer sb = new StringBuffer();
        String filepath = "layout3.html";
        String encoding = "GBK";

        try {
            // 读取文件将内容转成字符串
            FileInputStream fin = new FileInputStream(filepath);
            InputStreamReader reader = new InputStreamReader(fin, encoding);
            BufferedReader buffReader = new BufferedReader(reader);
            String strTmp;
            while ((strTmp = buffReader.readLine()) != null) {
                sb.append(strTmp);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        Cell cell = new Cell();
        LinkedHashMap map = cell.getTableInfo(sb);
        System.out.println(map.get("datajson"));
        System.out.println(map.get("pluginjson"));
    }
}
