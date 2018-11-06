import org.json.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.*;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.util.*;

public class Cell {

    public JSONObject getTable(String filepath, String encoding) {

        StringBuffer sb = new StringBuffer();
        try {
            // 读取文件将内容转成字符串
            File file = new File(filepath);
            FileInputStream fin = new FileInputStream(filepath);
            InputStreamReader reader = new InputStreamReader(fin, encoding);
            BufferedReader buffReader = new BufferedReader(reader);

            String strTmp = "";
            while ((strTmp = buffReader.readLine()) != null) {
                sb.append(strTmp);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        // 使用Jsoup解析html字符串
        Document document = Jsoup.parse(sb.toString());
        // 获取根节点
        ArrayList list = recursiveDFS(document);

        System.out.println(list);

        return null;
    }

    int m = 0; // 行
    int n = 0; // 列
    ArrayList cellList = new ArrayList();

    /**
     * 递归取出所有的节点信息，包括字段显示名和标题
     *
     * @param element
     * @return
     */
    public ArrayList recursiveDFS(Element element) {
        if (element.tag().toString().equals("tr") && element.child(0).tag().toString().equals("td")) {
            m++;
            n = 0;
        }
        if (element.tag().toString().equals("td")) {
            // 如果td里面是input框，则针对input框进行处理
            if (element.select("input").size() != 0) {
                for (int i = 0; i < element.select("input").size(); i++) {
                    HashMap inputMap = new HashMap();
                    String colspan = element.attr("colspan");
                    String name = element.select("input").get(i).attr("name");
                    String value = element.select("input").get(i).attr("value");
                    inputMap.put("id", m + "," + n);
                    inputMap.put("colspan", colspan.equals("") ? 1 : colspan);
                    inputMap.put("rowspan", 1);
                    inputMap.put("fieldid", name.substring(5));
                    inputMap.put("fieldtype", "text");
                    inputMap.put("etype", 3);
                    inputMap.put("evalue", value);
                    cellList.add(inputMap);
                    if (!colspan.equals("")) {
                        n += Integer.valueOf(colspan);
                    } else {
                        n++;
                    }
                    return cellList;
                }
                // 如果td里面是table，则认为其是明细表，进行特殊处理
            } else if (element.select("table").size() != 0) {

                /**
                 * code
                 */
                return cellList;
                // 明细表特殊处理
                // 否则就为普通的td
            } else {
                HashMap inputMap = new HashMap();
                String colspan = element.attr("colspan");
                String width = element.attr("width");
                String align = element.attr("align");
                String text = element.text();
                Element nextElement = element.nextElementSibling();
                if (nextElement != null && nextElement.select("input").size() != 0) {
                    for (int i = 0; i < nextElement.select("input").size(); i++) {
                        String name = nextElement.select("input").get(i).attr("name");
                        String value = nextElement.select("input").get(i).attr("value");
                        inputMap.put("evalue", value);
                        inputMap.put("fieldid", name.substring(5));
                    }
                }
                inputMap.put("id", m + "," + n);
                inputMap.put("colspan", colspan.equals("") ? 1 : colspan);
                inputMap.put("rowspan", 1);
                inputMap.put("width", width);
                inputMap.put("align", align);
                inputMap.put("etype", 2);
                inputMap.put("field", text);
                cellList.add(inputMap);
                if (!colspan.equals("")) {
                    n += Integer.valueOf(colspan);
                } else {
                    n++;
                }
                return cellList;
            }
        }
        for (Element child : element.children()) {
            recursiveDFS(child);
        }
        return cellList;
    }


    /**
     * 取出表单关联的js,css
     *
     * @param doc
     * @return
     */
    public HashMap getCssAndScript(Document doc) {
        HashMap map = new HashMap();
        int i = 0;
        for (Element e : doc.select("link")) {
            map.put(++i, e);
        }
        for (Element e : doc.select("script")) {
            map.put(++i, e);
        }
        return map;
    }
}
