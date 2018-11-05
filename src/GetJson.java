import org.json.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.util.HashMap;

public class GetJson {

    String filepath = "layout.html";

    public String get(String filepath, String encoding) {

        Attr attr = new Attr();

        try {
            // 读取文件将内容转成字符串
            File file = new File(filepath);
            FileInputStream fin = new FileInputStream(filepath);
            InputStreamReader reader = new InputStreamReader(fin, encoding);
            BufferedReader buffReader = new BufferedReader(reader);
            StringBuffer sb = new StringBuffer();
            String strTmp = "";
            while ((strTmp = buffReader.readLine()) != null) {
                sb.append(strTmp);
            }

            HashMap layout = new HashMap();

            // 使用Jsoup解析html字符串
            Document document = Jsoup.parse(sb.toString());

            Elements elements = document.getElementsByClass("ListStyle detailtable detailtableTopTable");

            for (int i = 0; i < elements.size(); i++) {
                HashMap cell = new HashMap();

                for (int j = 0; j < elements.get(i).select("td").size() / 2; j++) {

                    cell.put("id", i + "," + j);
                }
                layout.put("", cell);
            }



            System.out.println(new JSONObject(layout));

        } catch (Exception e) {
            e.printStackTrace();
        }
        return "";
    }

}
