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
import java.util.LinkedHashMap;

public class Cell {

    int mainRow = 0; // 主表行
    int mainCol = 0; // 主表列
    int detailRow = 0; // 明细表行
    int detailCol = 0; // 明细表列
    int detailCount = 0;
    ArrayList mainList = new ArrayList(); // 主表list
    LinkedHashMap eattr = new LinkedHashMap(); // 流程信息
    LinkedHashMap formula = new LinkedHashMap(); // 脚本和公式
    LinkedHashMap emaintable = new LinkedHashMap(); // 主表字段
    LinkedHashMap etables = new LinkedHashMap(); // 表单内容
    LinkedHashMap eformdesign = new LinkedHashMap();
    LinkedHashMap result = new LinkedHashMap();
    LinkedHashMap mainRowheads = new LinkedHashMap(); // 主表行统计
    LinkedHashMap mainColheads = new LinkedHashMap(); // 主表列统计
    LinkedHashMap detailRowheads = new LinkedHashMap(); // 明细表行统计
    LinkedHashMap detailColheads = new LinkedHashMap(); // 明细表列统计


    public JSONObject getTableInfo(String filepath, String encoding) {

        StringBuffer sb = new StringBuffer();
        try {
            // 读取文件将内容转成字符串
            File file = new File(filepath);
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

        Document document = Jsoup.parse(sb.toString());

        LinkedHashMap map = recursiveDFS(document);

        return new JSONObject(map);
    }


    /**
     * 递归取出所有的节点信息，包括字段显示名和标题
     *
     * @param element
     * @return
     */
    public LinkedHashMap recursiveDFS(Element element) {

        if (element.tag().toString().equals("tr") && element.child(0).tag().toString().equals("td")) {
            mainRowheads.put("row_" + mainRow, "30");
            emaintable.put("rowheads", mainRowheads);
            mainRow++; // 行数加一
            mainCol = 0; // 列归零
        }
        if (element.tag().toString().equals("td")) {
            // 如果td里面是table，则认为其是明细表，进行特殊处理
            if (element.select("table").size() != 0) {
                for (Element e : element.select("table")) {
                    // 如果该元素下面还有table元素，则继续循环，直到取到最底层的明细table，防止重复
                    if (e.select("table").size() > 1) {
                        continue;
                    }
                    LinkedHashMap map = detailTableAnalyse(e);
                    if (map.size() != 0) {
                        etables.put("detail_" + ++detailCount, map);
                    }
                }
                return etables;

                // 如果td里面是input框，则针对input框进行处理
            } else if (element.select("input").size() != 0) {
                for (int i = 0; i < element.select("input").size(); i++) {
                    LinkedHashMap inputMap = new LinkedHashMap();
                    String colspan = element.attr("colspan");
                    String width = element.attr("width");
                    String name = element.select("input").get(i).attr("name");
                    String value = element.select("input").get(i).attr("value");
                    inputMap.put("id", mainRow + "," + mainCol);
                    inputMap.put("colspan", colspan.equals("") ? "1" : colspan);
                    inputMap.put("rowspan", "1");
                    inputMap.put("width", width);
                    inputMap.put("field", name.substring(5));
                    inputMap.put("fieldtype", "text");
                    inputMap.put("etype", "3");
                    inputMap.put("evalue", value);

                    // 如果tr上有display:none属性，则记录下该显示属性
                    if (element.parent().attr("style").equals("display: none")) {
                        inputMap.put("display", "none");
                    }

                    mainList.add(inputMap);
                    emaintable.put("ec", mainList);
                    etables.put("emaintable", emaintable);
                }
                if (!element.attr("colspan").equals("")) {
                    mainCol += Integer.valueOf(element.attr("colspan"));
                } else {
                    mainCol++;
                }

                // 如果里面有strong标签，则作为标题处理
            } else if (element.select("strong").size() != 0) {
                LinkedHashMap map = new LinkedHashMap();
                ArrayList list = new ArrayList();
                map.put("id", mainRow + "," + 0);
                if (element.parent().select("font").size() != 0) {
                    String weight = "";
                    String color = element.parent().selectFirst("font").attr("color");
                    String size = element.parent().selectFirst("font").attr("size");
                    if (element.parent().selectFirst("font").attr("weight").equals("")) {
                        weight = element.parent().selectFirst("font").attr("size");
                    } else {
                        weight = element.parent().selectFirst("font").attr("weight");
                    }
                    map.put("color", color);
                    map.put("size", size);
                    map.put("weight", weight);
                }
                String colspan = element.attr("colspan");
                String width = element.attr("width");
                String align = element.attr("align");
                String text = element.text();
                map.put("colspan", colspan);
                map.put("width", width);
                map.put("align", align);
                map.put("text", text);
                list.add(map);
                emaintable.put("ec", list);
                mainRow++;
                return etables;

            } else { // 否则就为普通的td
                LinkedHashMap inputMap = new LinkedHashMap();
                String colspan = element.attr("colspan");
                String width = element.attr("width");
                String align = element.attr("align");
                String text = element.text();
                Element nextElement = element.nextElementSibling();
                if (nextElement != null && nextElement.select("input").size() != 0) {
                    for (int i = 0; i < nextElement.select("input").size(); i++) {
                        String name = nextElement.select("input").get(i).attr("name");
                        inputMap.put("field", name.substring(5));
                    }
                }
                inputMap.put("id", mainRow + "," + mainCol);
                inputMap.put("evalue", text);
                inputMap.put("colspan", colspan.equals("") ? "1" : colspan);
                inputMap.put("rowspan", "1");
                inputMap.put("width", width);
                inputMap.put("align", align);
                inputMap.put("etype", "2");
                if (element.parent().attr("style").equals("display: none")) {
                    inputMap.put("display", "none");
                }
                mainList.add(inputMap);
                emaintable.put("ec", mainList);
                etables.put("emaintable", emaintable);
                if (!colspan.equals("")) {
                    mainCol += Integer.valueOf(colspan);
                } else {
                    mainCol++;
                }
            }
        }
        for (Element child : element.children()) {
            recursiveDFS(child);
        }

//        getCssAndScript(element);
//        rs.execute(select * from workflow_nodehtmllayout where syspath like filepath)
        eattr.put("formname", "test");
        eattr.put("wfid", "9244");
        eattr.put("nodeid", "11995");
        eattr.put("formid", "-952");
        eattr.put("isbill", "-1");
        etables.put("emaintable", emaintable);
        eformdesign.put("eattr", eattr);
        eformdesign.put("etables", etables);
        eformdesign.put("formula", formula);
        result.put("eformdesign", eformdesign);
        return result;
    }


    /**
     * 单独处理明细表
     *
     * @param element
     * @return
     */
    public LinkedHashMap detailTableAnalyse(Element element) {
        ArrayList detailList = new ArrayList();
        LinkedHashMap detailMap = new LinkedHashMap();

        if (element.select("strong").size() != 0 && !(element.text().trim().equals(""))) {

            // 表单标题
            for (Element e : element.children()) {
                if ((e.text().trim().equals("")) || (e.text().trim().equals("&nbsp;"))) {
                    return etables;
                }
                if (e.tag().toString().equals("strong")) {
                    e = e.parent();

                    LinkedHashMap map = new LinkedHashMap();
                    map.put("id", mainRow + "," + 0);
                    if (e.parent().select("font").size() != 0) {
                        String color = e.attr("color");
                        String size = e.attr("size");
                        String weight = e.attr("weight");
                        map.put("color", color);
                        map.put("size", size);
                        map.put("weight", weight);
                    }
                    String colspan = e.attr("colspan");
                    String width = e.attr("width");
                    String align = e.attr("align");
                    String text = e.text();
                    map.put("colspan", colspan);
                    map.put("width", width);
                    map.put("align", align);
                    map.put("text", text);
                    mainList.add(map);
                    emaintable.put("ec", mainList);
                    mainRow++;
                    return etables;
                }
            }
        } else {
            for (Element detailtd : element.select("td")) {
                if (detailtd.select("input").size() != 0) {
                    for (int i = 0; i < detailtd.select("input").size(); i++) {
                        LinkedHashMap map = new LinkedHashMap();
                        String colspan = detailtd.attr("colspan");
                        String name = detailtd.select("input").get(i).attr("name");
                        String value = detailtd.select("input").get(i).attr("value");
                        map.put("id", detailRow + "," + detailCol);
                        if (detailRow == 1) map.put("attr", "content");
                        map.put("rowspan", 1);
                        map.put("field", name.substring(5));
                        map.put("fieldtype", "text");
                        map.put("etype", 3);
                        map.put("evalue", value);
                        detailList.add(map);
                        detailMap.put("ec", detailList);
                        detailCol++;
                    }
                } else if (!(detailtd.text().trim().equals("")) && !(detailtd.text().trim().equals("&nbsp;"))) {
                    // 否则就为普通的td
                    LinkedHashMap map = new LinkedHashMap();
                    String width = detailtd.attr("width");
                    String align = detailtd.attr("align");
                    String text = detailtd.text();
                    map.put("id", detailRow + "," + detailCol);
                    if (detailRow == 0) map.put("attr", "title");
                    map.put("rowspan", 1);
                    map.put("width", width);
                    map.put("align", align);
                    map.put("etype", 2);
                    map.put("field", text);
                    detailList.add(map);
                    detailMap.put("ec", detailList);
                    detailCol++;
                }
                if (detailtd.nextElementSibling() == null) {
                    if (detailRow == 1) {
                        detailRowheads.put("row_" + detailRow, "30");
                        detailRow = 0; // 明细表只有两行，一行为表头，一行为字段
                    } else {
                        detailRowheads.put("row_" + detailRow, "30");
                        detailRow++; // 下一行
                    }
                    detailCol = 0; // 列归零
                    detailMap.put("rowheads", detailRowheads);
                }
            }
            return detailMap;
        }
        for (Element e : element.children()) {
            detailTableAnalyse(e);
        }
        return detailMap;
    }


    /**
     * 取出表单关联的js,css
     *
     * @param element
     * @return
     */
    int count = 0;

    public LinkedHashMap getCssAndScript(Element element) {
        ArrayList list = new ArrayList();
        String elementTag = element.tag().toString();
        if (elementTag.equals("link") || elementTag.equals("script") || elementTag.equals("href")) {
            list.add(element);
            formula.put(list.size(), list);
        }
        return formula;
    }


    /**
     * 获取每一列的宽度
     * 总宽度除以标题栏的colspan值即为每一列的宽度
     *
     * @param element
     * @return
     */
//    public LinkedHashMap getEveryColWidth(Element element) {
//        Integer totalWidth = 0;
//        Integer colspan = 0;
//
//        for (Element e : element.getElementsByClass("table")) {
//
//            if (e.tag().toString().equals("table") && e.hasClass("table")) {
//                totalWidth = Integer.valueOf(element.attr("width"));
//                colspan = Integer.valueOf(element.selectFirst("td").attr("colspan"));
//                String perwidth = (totalWidth / colspan) + "";
//                for (int i = 0; i < colspan; i++) {
//                    detailColheads.put("col_", perwidth);
//                }
//            }
//        }
//        return detailColheads;
//    }
}

