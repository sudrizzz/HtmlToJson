import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.select.Elements;
import org.jsoup.select.NodeVisitor;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.HashMap;

public class Table {

    public String get(String filepath, String encoding) {
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

            // 使用Jsoup解析html字符串
//            Document document = Jsoup.parse(sb.toString());
            int i = sb.toString().indexOf("<table");
            int j = sb.toString().indexOf("</table>");

            Document doc = Jsoup.parse(sb.substring(i, j + "</table>".length()));

//            int tdi = sb.substring(i, j + "</table>".length()).lastIndexOf("<td");
//            int tdj = sb.substring(i, j + "</table>".length()).lastIndexOf("</td>");
//
//            Document doc2 = Jsoup.parse(sb.substring(i, j + "</table>".length()).substring(tdi, tdj + "</td>".length()));
//
//            System.out.println(doc2.text().trim());

            System.out.println(getTds(sb.toString()));

        } catch (Exception e) {
            e.printStackTrace();
        }
        return "";
    }


    // 取所有的table
    public Document getTables(String string) {
        int i = string.indexOf("<table");
        int j = string.indexOf("</table>");
        Document doc = Jsoup.parse(string.substring(i, j + "</table>".length()));

        return doc;
    }


    /**
     * 取所有的字段td
     *
     * @param string
     * @return
     */
    public HashMap<Integer, Document> getTds(String string) {
        HashMap<Integer, Document> map = new HashMap<>();
        Document doc = null;
        String primitiveStr = string;
        int tdEnd = string.indexOf("</td>");
        int i = 0;
        int j = 0;
        int k = 0;

        /**
         * 循环取最底层的td
         */
        while (true) {
            while (string.lastIndexOf("<td") != i) {
                i = string.lastIndexOf("<td");
                j = string.lastIndexOf("</td>");
                string = string.substring(string.indexOf("<td"), string.indexOf("</td>") + 6);
            }
            tdEnd = string.indexOf("</td>");
            doc = Jsoup.parse(string.substring(i, j + "</td>".length()));
            map.put(k++, doc);
            string = primitiveStr.substring(tdEnd, primitiveStr.length());
            i = string.indexOf("<td");
            j = string.indexOf("</td>");
            if (i == -1) {
                break;
            }
        }
        return map;
    }
}
