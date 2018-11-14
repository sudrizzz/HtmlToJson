import org.json.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import sun.awt.image.ImageWatched;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map;

public class Cell {

    int mainRow = 0; // 主表行
    int mainCol = 0; // 主表列
    int detailRow = 0; // 明细表行
    int detailCol = 0; // 明细表列
    int detailCount = 0;
    double dtTotalColWidth = 0.0;
    ArrayList mainList = new ArrayList(); // 主表list
    ArrayList detail; // 明细表list
    LinkedHashMap eattr = new LinkedHashMap<>(); // 流程信息
    LinkedHashMap formula = new LinkedHashMap(); // 脚本和公式
    LinkedHashMap emaintable = new LinkedHashMap(); // 主表字段
    LinkedHashMap etables = new LinkedHashMap(); // 表单内容
    LinkedHashMap<String, LinkedHashMap> eformdesign = new LinkedHashMap();
    LinkedHashMap mainRowheads = new LinkedHashMap(); // 主表行宽
    LinkedHashMap mainColheads = new LinkedHashMap(); // 主表列宽
    LinkedHashMap detailRowheads = new LinkedHashMap(); // 明细表行宽
    LinkedHashMap dataJsonMap = new LinkedHashMap();
    LinkedHashMap pluginMaintableData = new LinkedHashMap();
    LinkedHashMap pluginDetailtableData = new LinkedHashMap();
    LinkedHashMap pluginJsonMap = new LinkedHashMap();
    LinkedHashMap mainSheet = new LinkedHashMap();
    LinkedHashMap plugin = new LinkedHashMap();

    ParseSheet parseSheet = new ParseSheet();

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
        int totalWidth = Integer.valueOf(document.selectFirst("table[width]").attr("width"));
        int totalCols = Integer.valueOf(document.selectFirst("td[colspan]").attr("colspan"));
        for (int i = 0; i < totalCols; i++) {
            mainColheads.put("col_" + i, totalWidth / totalCols);
        }
        dtTotalColWidth = Double.valueOf(document.selectFirst("table[width]").attr("width"));

        eattr.put("formname", "htmltojson");
        eattr.put("wfid", "59729");
        eattr.put("nodeid", "78983");
        eattr.put("formid", "6");
        eattr.put("isbill", "0");
        emaintable.put("rowheads", mainRowheads);
        emaintable.put("colheads", mainColheads);
        emaintable.put("ec", mainList);
        etables.put("emaintable", emaintable);
        eformdesign.put("eattr", eattr);
        eformdesign.put("etables", etables);
        dataJsonMap.put("eformdesign", eformdesign);

        Map map = recursiveDFS(document);

        getPluginMap("main_sheet", new LinkedHashMap(), pluginMaintableData);

        System.out.println(new JSONObject(plugin));
        return new JSONObject(map);
    }


    /**
     * 递归取出所有的节点信息，包括字段显示名和标题
     *
     * @param element
     * @return
     */
    public LinkedHashMap recursiveDFS(Element element) {
        if (element.tag().toString().equals("tr")
                && element.child(0).tag().toString().equals("td")
                && element.select("table").size() == 0) {
            if (((ArrayList)emaintable.get("ec")).size() != 0) {
                mainRow++; // 行数加一
            }
            mainCol = 0; // 列归零
            mainRowheads.put("row_" + mainRow, "30");
            if (!mainRowheads.containsKey("row_0")) {
                mainRowheads.put("row_0", "30");
            }
        }
        if (element.tag().toString().equals("tr")) {
            LinkedHashMap map3 = new LinkedHashMap();
            for (int i = 0; i < element.select("td").size(); i++) {
                Element e = element.select("td").get(i);
                if (e.tag().toString().equals("td")) {
                    if (e.select("table").size() > 0) { // 如果td里面是table，则认为其是明细表，进行特殊处理
                        for (Element element1 : element.select("table")) { // 如果该元素下面还有table元素，则继续循环，直到取到最底层的明细table，防止重复
                            if (element1.select("table").size() > 1) {
                                continue;
                            }
                            detail = new ArrayList();
                            LinkedHashMap map = detailTableAnalyse(element1);
                            if (map.size() != 0) {
                                etables.put("detail_" + ++detailCount, map);
                                getPluginMap("detail_" + detailCount + "_sheet", new LinkedHashMap(), pluginDetailtableData);
                            }
                        }
                        return etables;
                        // 如果td里面是input框，则针对input框进行处理
                    } else if (e.select("input").size() > 0) {
                        for (int j = 0; j < e.select("input").size(); j++) {
                            CellAttr cellAttr = new CellAttr();
                            cellAttr.setRowid(mainRow);
                            cellAttr.setColid(mainCol);
                            cellAttr.setColspan(e.attr("colspan").equals("") ? 1 : Integer.valueOf(e.attr("colspan")));
                            cellAttr.setRowspan(e.attr("rowspan").equals("") ? 1 : Integer.valueOf(e.attr("rowspan")));
                            cellAttr.setEtype(3); // 2,字段名；3,表单内容
                            cellAttr.setFieldid(e.select("input").get(j).attr("name").substring(5));
                            cellAttr.setFieldattr(2); // 编辑
                            cellAttr.setFieldtype(e.select("input").get(j).attr("type")); // 单行文本框
                            cellAttr.setEvalue(e.select("input").get(j).attr("value"));
                            cellAttr.setHalign(1); // 左右居中
                            cellAttr.setValign(1); // 上下居中
                            cellAttr.setBtop_style(1);
                            cellAttr.setBtop_color("#90badd");
                            cellAttr.setBbottom_style(1);
                            cellAttr.setBbottom_color("#90badd");
                            cellAttr.setBleft_style(1);
                            cellAttr.setBleft_color("#90badd");
                            cellAttr.setBright_style(1);
                            cellAttr.setBright_color("#90badd");

                            LinkedHashMap map = new LinkedHashMap();
                            LinkedHashMap map2 = new LinkedHashMap();
                            parseSheet.buildDataEcMap(cellAttr, map);
                            parseSheet.buildPluginCellMap(cellAttr, map2);
                            map3.put("" + mainCol, map2);
                            mainList.add(map);
                            emaintable.put("ec", mainList);
                        }
                        if (!e.attr("colspan").equals("")) {
                            mainCol += Integer.valueOf(e.attr("colspan"));
                        } else {
                            mainCol++;
                        }
                        // 如果里面有strong标签，则作为明细标题处理
                    } else if (e.select("strong").size() > 0 && element.getElementsByClass("tr").size() == 0) {
                        int m = element.getElementsByClass("tr").size();
                        int n = e.select("strong").size();

                        CellAttr cellAttr = new CellAttr();
                        if (e.select("font").size() > 0) {
                            cellAttr.setFont_size(e.selectFirst("font").attr("size"));
                        } else {
                            cellAttr.setFont_size("12");
                        }
                        cellAttr.setRowid(mainRow);
                        cellAttr.setColid(mainCol);
                        cellAttr.setColspan(e.attr("colspan").equals("") ? 1 : Integer.valueOf(e.attr("colspan")));
                        cellAttr.setRowspan(e.attr("rowspan").equals("") ? 1 : Integer.valueOf(e.attr("rowspan")));
                        cellAttr.setEtype(1); // 1,文本；2,字段名；3,表单内容
                        cellAttr.setFieldattr(1); // 1编辑，2必填，3只读
                        cellAttr.setEvalue(e.text());
                        cellAttr.setHalign(1); // 左右居中
                        cellAttr.setValign(1); // 上下居中
                        cellAttr.setFont_color(e.attr("color"));
                        cellAttr.setBtop_style(1);
                        cellAttr.setBtop_color("#90badd");
                        cellAttr.setBbottom_style(1);
                        cellAttr.setBbottom_color("#90badd");
                        cellAttr.setBleft_style(1);
                        cellAttr.setBleft_color("#90badd");
                        cellAttr.setBright_style(1);
                        cellAttr.setBright_color("#90badd");

                        LinkedHashMap map = new LinkedHashMap();
                        LinkedHashMap map2 = new LinkedHashMap();
                        parseSheet.buildDataEcMap(cellAttr, map);
                        parseSheet.buildPluginCellMap(cellAttr, map2);
                        map3.put("" + mainCol, map2);
                        mainList.add(map);
                        emaintable.put("ec", mainList);
                        // 否则就为普通的td
                    } else {
                        CellAttr cellAttr = new CellAttr();
                        String colspan = e.attr("colspan");
                        String rowspan = e.attr("rowspan");
                        Element nextElement = e.nextElementSibling();
                        if (nextElement != null && nextElement.select("input").size() > 0) {
                            for (int j = 0; j < nextElement.select("input").size(); j++) {
                                cellAttr.setFieldid(nextElement.select("input").get(j).attr("name").substring(5));
                            }
                        }
                        cellAttr.setRowid(mainRow);
                        cellAttr.setColid(mainCol);
                        cellAttr.setColspan(colspan.equals("") ? 1 : Integer.valueOf(colspan));
                        cellAttr.setRowspan(rowspan.equals("") ? 1 : Integer.valueOf(rowspan));
                        cellAttr.setEtype(2); // 2,字段名；3,表单内容
                        cellAttr.setEvalue(e.text());
                        cellAttr.setHalign(1); // 左右居中
                        cellAttr.setValign(1); // 上下居中
                        cellAttr.setBtop_style(1);
                        cellAttr.setBtop_color("#90badd");
                        cellAttr.setBbottom_style(1);
                        cellAttr.setBbottom_color("#90badd");
                        cellAttr.setBleft_style(1);
                        cellAttr.setBleft_color("#90badd");
                        cellAttr.setBright_style(1);
                        cellAttr.setBright_color("#90badd");

                        LinkedHashMap map = new LinkedHashMap();
                        LinkedHashMap map2 = new LinkedHashMap();
                        parseSheet.buildDataEcMap(cellAttr, map);
                        parseSheet.buildPluginCellMap(cellAttr, map2);
                        map3.put("" + mainCol, map2);
                        mainList.add(map);
                        emaintable.put("ec", mainList);

                        if (!colspan.equals("")) {
                            mainCol += Integer.valueOf(colspan);
                        } else {
                            mainCol++;
                        }
                    }
                }
            }
            pluginMaintableData.put(mainRow, map3);
        }
        for (Element child : element.children()) {
            recursiveDFS(child);
        }
        return dataJsonMap;
    }


    /**
     * 单独处理明细表
     *
     * @param element
     * @return
     */
    private LinkedHashMap detailTableAnalyse(Element element) {
        LinkedHashMap detailMap = new LinkedHashMap();
        if (element.select("strong").size() > 0 && !(element.text().trim().equals(""))) {
            for (Element e : element.children()) { // 表单标题
                if ((e.text().trim().equals("")) || (e.text().trim().equals("&nbsp;"))) {
                    return etables;
                }
                if (e.tag().toString().equals("strong")) {
                    e = e.parent();
                    Element hasColspanEle = e;
                    CellAttr cellAttr = new CellAttr();
                    cellAttr.setRowid(mainRow);
                    cellAttr.setColid(0);
                    if (e.attr("colspan").equals("")) {
                        while (hasColspanEle.attr("colspan").equals("")) {
                            hasColspanEle = hasColspanEle.parent();
                        }
                        cellAttr.setColspan(Integer.valueOf(hasColspanEle.attr("colspan")));
                        cellAttr.setRowspan(e.attr("rowspan").equals("") ? 1 : Integer.valueOf(e.attr("rowspan")));
                    } else {
                        cellAttr.setColspan(e.attr("colspan").equals("") ? 1 : Integer.valueOf(e.attr("colspan")));
                        cellAttr.setRowspan(e.attr("rowspan").equals("") ? 1 : Integer.valueOf(e.attr("rowspan")));
                    }
                    cellAttr.setEtype(1); // 2,字段名；3,表单内容
                    cellAttr.setFieldattr(3); // 1编辑，2必填，3只读
                    cellAttr.setEvalue(e.text());
                    cellAttr.setHalign(1); // 左右居中
                    cellAttr.setValign(1); // 上下居中
                    cellAttr.setFont_color(e.attr("color"));
                    cellAttr.setBtop_style(1);
                    cellAttr.setBtop_color("#90badd");
                    cellAttr.setBbottom_style(1);
                    cellAttr.setBbottom_color("#90badd");
                    cellAttr.setBleft_style(1);
                    cellAttr.setBleft_color("#90badd");
                    cellAttr.setBright_style(1);
                    cellAttr.setBright_color("#90badd");

                    LinkedHashMap map = new LinkedHashMap();
                    LinkedHashMap map2 = new LinkedHashMap();
                    parseSheet.buildDataEcMap(cellAttr, map);
                    parseSheet.buildPluginCellMap(cellAttr, map2);
                    mainList.add(map);
                    emaintable.put("ec", mainList);
                }
            }
        } else {
            for (Element e : element.select("tr")) {
                LinkedHashMap map3 = new LinkedHashMap();
                for (Element detailtd : e.select("td")) {
                    if (detailtd.select("input").size() != 0) {
                        for (int i = 0; i < detailtd.select("input").size(); i++) {
                            CellAttr cellAttr = new CellAttr();
                            cellAttr.setRowid(detailRow);
                            cellAttr.setColid(detailCol);
                            cellAttr.setColspan(detailtd.attr("colspan").equals("") ? 1 : Integer.valueOf(detailtd.attr("colspan")));
                            cellAttr.setRowspan(detailtd.attr("rowspan").equals("") ? 1 : Integer.valueOf(detailtd.attr("rowspan")));
                            cellAttr.setEtype(3); // 2,字段名；3,表单内容
                            cellAttr.setFieldid(detailtd.select("input").get(i).attr("name").substring(5));
                            cellAttr.setFieldattr(1); // 编辑
                            cellAttr.setFieldtype("text"); // 单行文本框
                            cellAttr.setEvalue(detailtd.select("input").get(i).attr("value"));
                            cellAttr.setHalign(1); // 左右居中
                            cellAttr.setValign(1); // 上下居中
                            cellAttr.setBtop_style(1);
                            cellAttr.setBtop_color("#90badd");
                            cellAttr.setBbottom_style(1);
                            cellAttr.setBbottom_color("#90badd");
                            cellAttr.setBleft_style(1);
                            cellAttr.setBleft_color("#90badd");
                            cellAttr.setBright_style(1);
                            cellAttr.setBright_color("#90badd");

                            LinkedHashMap map = new LinkedHashMap();
                            LinkedHashMap map2 = new LinkedHashMap();
                            parseSheet.buildDataEcMap(cellAttr, map);
                            parseSheet.buildPluginCellMap(cellAttr, map2);
                            map3.put("" + detailCol, map2);
                            pluginDetailtableData.put(detailRow, map3);
                            detail.add(map);
                            detailMap.put("ec", detail);
                            detailCol++;
                        }
                        // 否则就为普通的td
                    } else {
                        CellAttr cellAttr = new CellAttr();
                        cellAttr.setEvalue(detailtd.text());
                        cellAttr.setRowid(detailRow);
                        cellAttr.setColid(detailCol);
                        cellAttr.setColspan(1);
                        cellAttr.setRowspan(1);
                        cellAttr.setEtype(2); // 2,字段名；3,表单内容
                        cellAttr.setFieldattr(3); // 编辑
                        cellAttr.setEvalue(detailtd.text());
                        cellAttr.setHalign(1); // 左右居中
                        cellAttr.setValign(1); // 上下居中
                        cellAttr.setBtop_style(1);
                        cellAttr.setBtop_color("#90badd");
                        cellAttr.setBbottom_style(1);
                        cellAttr.setBbottom_color("#90badd");
                        cellAttr.setBleft_style(1);
                        cellAttr.setBleft_color("#90badd");
                        cellAttr.setBright_style(1);
                        cellAttr.setBright_color("#90badd");

                        LinkedHashMap map = new LinkedHashMap();
                        LinkedHashMap map2 = new LinkedHashMap();
                        parseSheet.buildDataEcMap(cellAttr, map);
                        parseSheet.buildPluginCellMap(cellAttr, map2);
                        map3.put("" + detailCol, map2);
                        pluginDetailtableData.put(detailRow, map3);
                        detail.add(map);
                        detailMap.put("ec", detail);
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
        LinkedHashMap linkMap = new LinkedHashMap();
        LinkedHashMap scriptMap = new LinkedHashMap();
        int i = 0;

        if (element.select("link").size() != 0
                || element.select("script").size() != 0) {
            for (Element e : element.select("link")) {
                linkMap.put(i++, e.toString());
            }
            i = 0;
            for (Element e : element.select("script")) {
                scriptMap.put(i++, e.toString());
            }
        }
        formula.put("link", linkMap);
        formula.put("script", scriptMap);
        return formula;
    }


//    /**
//     * 获取每一列的宽度
//     * 总宽度除以标题栏的colspan值即为每一列的宽度
//     *
//     * @param element
//     * @return
//     */
//    public LinkedHashMap getEveryColWidth(Element element) {
//        LinkedHashMap detailColheads = new LinkedHashMap(); // 明细表列宽
//        Integer totalWidth = 0;
//        Integer colspan = 0;
//
//        for (Element e : element.getElementsByClass("table")) {
//
//            if (e.tag().toString().equals("table") && e.hasClass("table")) {
//                totalWidth = Integer.valueOf(element.attr("width"));
//                colspan = Integer.valueOf(element.selectFirst("td").attr("colspan"));
//                String perColWidth = (totalWidth / colspan) + "";
//                for (int i = 0; i < colspan; i++) {
//                    detailColheads.put("col_", perColWidth);
//                }
//            }
//        }
//        return detailColheads;
//    }



    private void getPluginMap(String str, Map pluginMap, Map data) {
        pluginMap.put("version", "2.0");
        pluginMap.put("tabStripVisible", false);
        pluginMap.put("canUserEditFormula", false);
        pluginMap.put("allowUndo", false);
        pluginMap.put("allowDragDrop", false);
        pluginMap.put("allowDragFill", false);
        pluginMap.put("grayAreaBackColor", "white");
        Map<String, Object> sheets = new LinkedHashMap<String, Object>();
        pluginMap.put("sheets", sheets);
        Map<String, Object> Sheet1 = new LinkedHashMap<String, Object>();
        sheets.put("Sheet1", Sheet1);

        Sheet1.put("name", "Sheet1");
        Map<String, Object> defaults = new LinkedHashMap<String, Object>();
        defaults.put("rowHeight", 30);
        defaults.put("colWidth", 62);
        defaults.put("rowHeaderColWidth", 40);
        defaults.put("colHeaderRowHeight", 20);
        Sheet1.put("defaults", defaults);
        Map<String, Object> selections = new LinkedHashMap<String, Object>();
        Map<String, Object> selections_0 = new LinkedHashMap<String, Object>();
        selections.put("0", selections_0);
        selections_0.put("row", 0);
        selections_0.put("rowCount", 1);
        selections_0.put("col", 0);
        selections_0.put("colCount", 4);
        Sheet1.put("selections", selections);
        Sheet1.put("activeRow", 0);
        Sheet1.put("activeCol", 0);
        Map<String, Object> gridline = new LinkedHashMap<String, Object>();
        gridline.put("color", "#D0D7E5");
        gridline.put("showVerticalGridline", true);
        gridline.put("showHorizontalGridline", true);
        Sheet1.put("gridline", gridline);
        Sheet1.put("allowDragDrop", false);
        Sheet1.put("allowDragFill", false);

        Map<String, Object> rowHeaderData = new LinkedHashMap<String, Object>();
        Map<String, Object> colHeaderData = new LinkedHashMap<String, Object>();
        Map<String, Object> defaultDataNode = new LinkedHashMap<String, Object>();
        Map<String, Object> defaultDataNode_style = new LinkedHashMap<String, Object>();
        defaultDataNode_style.put("foreColor", "black");
        defaultDataNode.put("style", defaultDataNode_style);
        rowHeaderData.put("rowCount", parseSheet.getRownum() + 5);
        rowHeaderData.put("defaultDataNode", defaultDataNode);
        colHeaderData.put("colCount", parseSheet.getColnum() + 3);
        colHeaderData.put("defaultDataNode", defaultDataNode);
        Sheet1.put("rowHeaderData", rowHeaderData);
        Sheet1.put("colHeaderData", colHeaderData);

        Map dataTable = new LinkedHashMap();
        dataTable.put("dataTable", data);
        Sheet1.put("rows", parseSheet.getPlugin_rows());
        Sheet1.put("columns", getPluginColumns(data));
//        Sheet1.put("spans", getPlugin_combine());
        Sheet1.put("data", dataTable);


        Map<String, Object> rowRangeGroup = new LinkedHashMap<String, Object>();
        rowRangeGroup.put("itemsCount", parseSheet.getRownum() + 5);
        Map<String, Object> colRangeGroup = new LinkedHashMap<String, Object>();
        colRangeGroup.put("itemsCount", parseSheet.getColnum() + 3);
        Sheet1.put("rowRangeGroup", rowRangeGroup);
        Sheet1.put("colRangeGroup", colRangeGroup);

        plugin.put(str, pluginMap);
    }


    private ArrayList getPluginColumns(Map map) {
        ArrayList list = new ArrayList();
        for (int i = 0; i < ((LinkedHashMap)map.get(detailRow + 1)).size(); i++) {
            LinkedHashMap<String,Object> plugin_column = new LinkedHashMap<String,Object>();
            plugin_column.put("size", "120");
            plugin_column.put("dirty", "true");
            list.add(plugin_column);
        }
        return list;
    }


    private ArrayList getPlugin_combine(){
        ArrayList list = new ArrayList();
        LinkedHashMap<String,Object> plugin_column = new LinkedHashMap<String,Object>();
        plugin_column.put("row", "0");
        plugin_column.put("rowCount", "1");
        plugin_column.put("col", "0");
        plugin_column.put("colCount", "6");
        list.add(plugin_column);
        return list;
    }
}

