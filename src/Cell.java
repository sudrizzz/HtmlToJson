import org.json.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.HashMap;

public class Cell {

    public JSONObject getTableInfo(String filepath, String encoding) {

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
        Document document = Jsoup.parse(sb.toString());
        // 获取根节点
        HashMap map = recursiveDFS(document);

        System.out.println(new JSONObject(map));

        return null;
    }

    int mainRow = 0; // 主表行
    int mainCol = 0; // 主表列
    ArrayList cellList = new ArrayList(); // 主表list
    HashMap emaintable = new HashMap();
    HashMap detail = new HashMap();
    HashMap etables = new HashMap();

    /**
     * 递归取出所有的节点信息，包括字段显示名和标题
     *
     * @param element
     * @return
     */
    public HashMap recursiveDFS(Element element) {
        if (element.tag().toString().equals("tr") && element.child(0).tag().toString().equals("td")) {
            mainRow++; // 下一行
            mainCol = 0; // 列归零
        }
        if (element.tag().toString().equals("td")) {
            // 如果td里面是table，则认为其是明细表，进行特殊处理
            if (element.select("table").size() != 0) {
                for (Element e : element.select("table")) {
                    System.out.println(detailTableAnalyse(e));
                }
                return etables;
                // 如果td里面是input框，则针对input框进行处理
            } else if (element.select("input").size() != 0) {
                for (int i = 0; i < element.select("input").size(); i++) {
                    HashMap inputMap = new HashMap();
                    String colspan = element.attr("colspan");
                    String name = element.select("input").get(i).attr("name");
                    String value = element.select("input").get(i).attr("value");
                    inputMap.put("id", mainRow + "," + mainCol);
                    inputMap.put("colspan", colspan.equals("") ? "1" : colspan);
                    inputMap.put("rowspan", "1");
                    inputMap.put("field", name.substring(5));
                    inputMap.put("fieldtype", "text");
                    inputMap.put("etype", "3");
                    inputMap.put("evalue", value);
                    cellList.add(inputMap);
                    emaintable.put("ec", cellList);
                    etables.put("emaintable", emaintable);
                    if (!colspan.equals("")) {
                        mainCol += Integer.valueOf(colspan);
                    } else {
                        mainCol++;
                    }
                    return etables;
                }
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
                        inputMap.put("field", name.substring(5));
                    }
                }
                inputMap.put("id", mainRow + "," + mainCol);
                inputMap.put("colspan", colspan.equals("") ? "1" : colspan);
                inputMap.put("rowspan", "1");
                inputMap.put("width", width);
                inputMap.put("align", align);
                inputMap.put("etype", "2");
                cellList.add(inputMap);
                emaintable.put("ec", cellList);
                etables.put("emaintable", emaintable);
                if (!colspan.equals("")) {
                    mainCol += Integer.valueOf(colspan);
                } else {
                    mainCol++;
                }
            }
        }
        for (Element child : element.children()) {
            emaintable.put("ec", cellList);
            etables.put("emaintable", emaintable);
            recursiveDFS(child);
        }
        return etables;
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


    int detailRow = 0; // 主表行
    int detailCol = 0; // 主表列
    ArrayList detailList = new ArrayList();

    /**
     * 单独处理明细表
     *
     * @param element
     * @return
     */
    public ArrayList detailTableAnalyse(Element element) {
        if (element.tag().toString().equals("tr") && element.child(0).tag().toString().equals("td")) {
            detailRow++; // 下一行
            detailCol = 0; // 列归零
        }
        if (element.select("strong").size() != 0 && !(element.text().trim().equals(""))) {
            // 表单标题
            for (Element e : element.select("strong").parents()) {
                if ((e.text().trim().equals("")) || (e.text().trim().equals("&nbsp;"))) {
                    return detailList;
                }
                HashMap detailMap = new HashMap();
                String color = e.attr("color");
                String size = e.attr("size");
                String weight = e.attr("weight");
                String colspan = e.attr("colspan");
                String width = e.attr("width");
                String align = e.attr("align");
                String text = e.text();
                detailMap.put("id", color);
                detailMap.put("color", color);
                detailMap.put("size", size);
                detailMap.put("weight", weight);
                detailMap.put("colspan", colspan);
                detailMap.put("width", width);
                detailMap.put("align", align);
                detailMap.put("text", text);
                detailList.add(detailMap);
                detailCol++;
                return detailList;
            }
        } else {
            for (Element detailtd : element.select("td")) {
                if (detailtd.select("input").size() != 0) {
                    for (int i = 0; i < detailtd.select("input").size(); i++) {
                        HashMap detailMap = new HashMap();
                        String colspan = detailtd.attr("colspan");
                        String name = detailtd.select("input").get(i).attr("name");
                        String value = detailtd.select("input").get(i).attr("value");
                        detailMap.put("id", detailRow + "," + detailCol);
                        detailMap.put("rowspan", 1);
                        detailMap.put("field", name.substring(5));
                        detailMap.put("fieldtype", "text");
                        detailMap.put("etype", 3);
                        detailMap.put("evalue", value);
                        detailList.add(detailMap);
                    }
                } else if (!(detailtd.text().trim().equals("")) && !(detailtd.text().trim().equals("&nbsp;"))) {
                    // 否则就为普通的td
                    HashMap detailMap = new HashMap();
                    String width = detailtd.attr("width");
                    String align = detailtd.attr("align");
                    String text = detailtd.text();
                    detailMap.put("id", detailRow + "," + detailCol);
                    detailMap.put("rowspan", 1);
                    detailMap.put("width", width);
                    detailMap.put("align", align);
                    detailMap.put("etype", 2);
                    detailMap.put("field", text);
                    detailList.add(detailMap);
                    detailCol++;
                }
            }
            return detailList;
        }
        for (Element e : element.select("td")) {
            detailTableAnalyse(e);
        }
        return detailList;
    }
}

