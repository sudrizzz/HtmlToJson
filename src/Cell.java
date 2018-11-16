import net.sf.json.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.LinkedHashMap;

public class Cell {

    int mainRow = 0; // 主表行
    int mainCol = 0; // 主表列
    int detailRow = 0; // 明细表行
    int detailCol = 0; // 明细表列
    int detailCount = 0; // 明细计数
    int totalCols = 0; // 主表单元格数
    ArrayList mainList = new ArrayList(); // 主表list
    ArrayList detail; // 明细表list
    ArrayList mainSpans = new ArrayList(); // 主表合并单元格信息
    ArrayList detailSpans = new ArrayList(); // 明细合并单元格信息
    LinkedHashMap eattr = new LinkedHashMap<>(); // 流程信息
    //    LinkedHashMap formula = new LinkedHashMap(); // 脚本和公式
    LinkedHashMap emaintable = new LinkedHashMap(); // 主表字段
    LinkedHashMap etables = new LinkedHashMap(); // 表单内容
    LinkedHashMap<String, Object> plugin = new LinkedHashMap();
    LinkedHashMap<String, Object> eformdesign = new LinkedHashMap();
    LinkedHashMap<String, Object> mainRowheads = new LinkedHashMap(); // 主表行宽
    LinkedHashMap<String, Object> mainColheads = new LinkedHashMap(); // 主表列宽
    LinkedHashMap<String, Object> detailRowheads = new LinkedHashMap(); // 明细表行高
    LinkedHashMap<String, Object> detailColheads = new LinkedHashMap(); // 明细表列宽
    LinkedHashMap<String, Object> dataJsonMap = new LinkedHashMap(); // datajson
    LinkedHashMap<String, Object> pluginMaintableData = new LinkedHashMap(); // pluginjson 主表
    LinkedHashMap<String, Object> pluginDetailtableData = new LinkedHashMap(); // pluginjson 明细表

    ParseSheet parseSheet = new ParseSheet();

    public LinkedHashMap getTableInfo(String filepath, String encoding) {
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
        totalCols = Integer.valueOf(document.selectFirst("td[colspan]").attr("colspan"));
        for (int i = 0; i < totalCols; i++) {
            mainColheads.put("col_" + i, totalWidth / totalCols);
        }

        eattr.put("formname", "htmltojson");
        eattr.put("wfid", "59729");
        eattr.put("nodeid", "78983");
        eattr.put("formid", "6");
        eattr.put("isbill", "1");
        emaintable.put("rowheads", mainRowheads);
        emaintable.put("colheads", mainColheads);
        emaintable.put("ec", mainList);
        etables.put("emaintable", emaintable);
        eformdesign.put("eattr", eattr);
        eformdesign.put("etables", etables);
        dataJsonMap.put("eformdesign", eformdesign);
        LinkedHashMap datajsonMap = recursiveDFS(document);
        getPluginMap("main_sheet", new LinkedHashMap(), pluginMaintableData);

        LinkedHashMap jsonMap = new LinkedHashMap();
        jsonMap.put("datajson", datajsonMap);
        jsonMap.put("pluginjson", JSONObject.fromObject(plugin));
        return jsonMap;
    }


    /**
     * 递归取出所有的节点信息，包括字段显示名和标题
     */
    public LinkedHashMap recursiveDFS(Element element) {
        if (element.tag().toString().equals("tr")
                && element.child(0).tag().toString().equals("td")
                && element.select("table").size() == 0) {
            if (((ArrayList) emaintable.get("ec")).size() != 0) {
                mainRow++; // 行数加一
            }
            mainCol = 0; // 列归零
            mainRowheads.put("row_" + mainRow, "30");
            if (!mainRowheads.containsKey("row_0")) {
                mainRowheads.put("row_0", "30");
            }
        }
        if (element.tag().toString().equals("tr")) {
            LinkedHashMap pluginData = new LinkedHashMap();
            for (int i = 0; i < element.select("td").size(); i++) {
                Element e = element.select("td").get(i);
                if (e.tag().toString().equals("td")) {
                    if (e.select("table").size() > 0) { // 如果td里面是table，则认为其是明细表，进行特殊处理
                        for (Element element1 : element.select("table")) {
                            // 如果该元素下面还有table元素，则继续循环，直到取到最底层的明细table，防止重复
                            if (element1.select("table").size() > 1) {
                                continue;
                            }
                            detail = new ArrayList();
                            LinkedHashMap map = detailTableAnalyse(element1);
                            if (map.size() != 0) {
                                etables.put("detail_" + ++detailCount, map);
                                getPluginMap("detail_" + detailCount + "_sheet", new LinkedHashMap(), pluginDetailtableData);
                                getdetailTableCell("detail_" + detailCount);

                                CellAttr cellAttr = new CellAttr();
                                getMainTableSpans(cellAttr, e);
                                cellAttr.setRowid(mainRow);
                                cellAttr.setColid(mainCol);
                                cellAttr.setEvalue("明细表" + detailCount);
                                cellAttr.setValign(1); // 上下居中
                                cellAttr.setBackground_color("#e7f3fc");
                                cellAttr.setBtop_style(1);
                                cellAttr.setBtop_color("#90badd");
                                cellAttr.setBbottom_style(1);
                                cellAttr.setBbottom_color("#90badd");
                                cellAttr.setBleft_style(1);
                                cellAttr.setBleft_color("#90badd");
                                cellAttr.setBright_style(1);
                                cellAttr.setBright_color("#90badd");

                                LinkedHashMap map2 = new LinkedHashMap();
                                parseSheet.buildPluginCellMap(cellAttr, map2);
                                pluginData.put("" + mainCol, map2);
                                pluginMaintableData.put(mainRow + "", pluginData);
                                detailSpans = new ArrayList(); // 重置明细表的列合并信息
                            }
                        }
                        return etables;
                        // 如果td里面是input框，则针对input框进行处理
                    } else if (e.select("input").size() > 0) {
                        for (int j = 0; j < e.select("input").size(); j++) {
                            CellAttr cellAttr = new CellAttr();
                            String attr = e.select("input").get(j).val();
                            if (attr.contains("序号")) {
                                cellAttr.setEtype(22); // 22,序号
                            } else {
                                cellAttr.setEtype(3); // 2,字段名；3,表单内容
                            }
                            if (attr.contains("[") && attr.contains("]")) {
                                cellAttr.setEvalue(attr.substring(attr.indexOf("]")));
                            } else {
                                cellAttr.setEvalue(attr);
                            }
                            getMainTableSpans(cellAttr, e);
                            cellAttr.setRowid(mainRow);
                            cellAttr.setColid(mainCol);
                            cellAttr.setFieldid(e.select("input").get(j).attr("name").substring(5));
                            cellAttr.setFieldattr(1); // 1,编辑；2,必填；3,只读
                            cellAttr.setFieldtype(e.select("input").get(j).attr("type")); // 单行文本框
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
                            pluginData.put("" + mainCol, map2);
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
                        CellAttr cellAttr = new CellAttr();
                        if (e.select("font").size() > 0) {
                            cellAttr.setFont_size(e.selectFirst("font").attr("size"));
                        } else {
                            cellAttr.setFont_size("12pt");
                        }
                        cellAttr.setBold(true);
                        getMainTableSpans(cellAttr, e);
                        cellAttr.setRowid(mainRow);
                        cellAttr.setColid(mainCol);
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
                        pluginData.put("" + mainCol, map2);
                        mainList.add(map);
                        emaintable.put("ec", mainList);
                    } else if (e.select("a").size() > 0) { // 超链接
                        CellAttr cellAttr = new CellAttr();
                        getMainTableSpans(cellAttr, e);
                        cellAttr.setRowid(mainRow);
                        cellAttr.setColid(mainCol);
                        cellAttr.setEtype(11); // 11,超链接
                        cellAttr.setEvalue(e.text());
                        cellAttr.setValign(1); // 上下居中
                        cellAttr.setFont_color("#0000ee");
                        cellAttr.setBtop_style(1);
                        cellAttr.setUnderline(true);
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
                        pluginData.put("" + mainCol, map2);
                        mainList.add(map);
                        emaintable.put("ec", mainList);
                    } else { // 否则就为普通的td
                        CellAttr cellAttr = new CellAttr();
                        getMainTableSpans(cellAttr, e);
                        String colspan = e.attr("colspan");
                        Element nextElement = e.nextElementSibling();
                        if (nextElement != null && nextElement.select("input").size() > 0) {
                            for (int j = 0; j < nextElement.select("input").size(); j++) {
                                cellAttr.setFieldid(nextElement.select("input").get(j).attr("name").substring(5));
                            }
                        }
                        cellAttr.setRowid(mainRow);
                        cellAttr.setColid(mainCol);
                        cellAttr.setEtype(2); // 2,字段名；3,表单内容
                        cellAttr.setEvalue(e.text());
                        cellAttr.setHalign(1); // 左右居中
                        cellAttr.setValign(1); // 上下居中
                        cellAttr.setBackground_color("#e7f3fc");
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
                        pluginData.put("" + mainCol, map2);
                        mainList.add(map);
                        emaintable.put("ec", mainList);

                        if (!colspan.equals("")) {
                            mainCol += Integer.valueOf(colspan);
                            if (mainCol >= totalCols) {
                                mainCol = 0;
                            }
                        } else {
                            mainCol++;
                        }
                    }
                }
            }
            pluginMaintableData.put(mainRow + "", pluginData);
        }
        for (Element child : element.children()) {
            recursiveDFS(child);
        }
        return dataJsonMap;
    }


    /**
     * 单独处理明细表
     */
    private LinkedHashMap detailTableAnalyse(Element element) {
        LinkedHashMap detailMap = new LinkedHashMap();
        if (element.select("strong").size() > 0 && !(element.text().trim().equals(""))) {
            LinkedHashMap pluginData = new LinkedHashMap();
            for (Element e : element.children()) { // 表单标题
                if ((e.text().trim().equals("")) || (e.text().trim().equals("&nbsp;"))) {
                    return etables;
                }
                if (e.tag().toString().equals("strong")) {
                    e = e.parent();
                    Element colspanElement = e;
                    CellAttr cellAttr = new CellAttr();
                    cellAttr.setRowid(mainRow);
                    cellAttr.setColid(0);
                    if (e.attr("colspan").equals("")) {
                        while (colspanElement.attr("colspan").equals("")) {
                            colspanElement = colspanElement.parent();
                        }
                        cellAttr.setColspan(Integer.valueOf(colspanElement.attr("colspan")));
                        getMergedCellsInfo(mainRow, mainCol, Integer.valueOf(colspanElement.attr("colspan")), "colspan");
                    } else {
                        cellAttr.setColspan(Integer.valueOf(e.attr("colspan")));
                        getMergedCellsInfo(mainRow, mainCol, Integer.valueOf(e.attr("colspan")), "colspan");
                    }
                    cellAttr.setFont_size("18pt");
                    cellAttr.setBold(true);
                    cellAttr.setRowspan(e.attr("rowspan").equals("") ? 1 : Integer.valueOf(e.attr("rowspan")));
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
                    pluginData.put("" + mainCol, map2);
                    mainList.add(map);
                    pluginMaintableData.put(mainRow + "", pluginData);
                    emaintable.put("ec", mainList);
                }
            }
        } else {
            pluginDetailtableData = new LinkedHashMap<>();
            for (Element e : element.select("tr")) {
                LinkedHashMap pluginData = new LinkedHashMap();
                int maxColumn = 0;
                for (Element detailtd : e.select("td")) {
                    if (detailtd.select("input").size() != 0) {
                        if (detailCol == 0) {
                            CellAttr cellAttr = new CellAttr();
                            cellAttr.setRowid(detailRow);
                            cellAttr.setColid(detailCol);
                            cellAttr.setEtype(21);
                            cellAttr.setEvalue("选中");
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
                            detail.add(map);
                            detailMap.put("ec", detail);
                            parseSheet.buildDataEcMap(cellAttr, map);
                            parseSheet.buildPluginCellMap(cellAttr, map2);
                            detailColheads.put("col_" + detailCol, "50");
                            pluginData.put("" + detailCol, map2);
                            pluginDetailtableData.put(detailRow + "", pluginData);
                            detailCol++;
                        }
                        for (int i = 0; i < detailtd.select("input").size(); i++) {
                            CellAttr cellAttr = new CellAttr();
                            String attr = detailtd.select("input").get(i).val();
                            if (attr.contains("序号")) {
                                cellAttr.setEtype(22); // 22,序号
                            } else {
                                cellAttr.setEtype(3); // 2,字段名；3,表单内容
                            }
                            cellAttr.setRowid(detailRow);
                            cellAttr.setColid(detailCol);
                            cellAttr.setColspan(1);
                            cellAttr.setRowspan(1);
                            cellAttr.setFieldid(detailtd.select("input").get(i).attr("name").substring(5));
                            cellAttr.setFieldtype("text"); // 单行文本框
                            cellAttr.setEvalue(attr.substring(4));
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
                            pluginData.put("" + detailCol, map2);
                            pluginDetailtableData.put(detailRow + "", pluginData);
                            detail.add(map);
                            detailMap.put("ec", detail);
                            detailCol++;
                        }
                        // 否则就为普通的td
                    } else {
                        if (detailMap.size() == 0) {
                            for (int i = 0; i < e.select("td").size() + 1; i++) {
                                CellAttr cellAttr = new CellAttr();
                                // 明细空一行的最后一格显示添加和删除按钮
                                if (i == e.select("td").size()) {
                                    cellAttr.setEtype(10);
                                }
                                cellAttr.setRowid(detailRow);
                                cellAttr.setColid(detailCol);
                                LinkedHashMap map = new LinkedHashMap();
                                detail.add(map);
                                detailMap.put("ec", detail);
                                parseSheet.buildDataEcMap(cellAttr, map);
                                detailCol++;
                            }
                            detailCol = 0;
                            for (int i = 0; i < e.select("td").size() + 1; i++) {
                                CellAttr cellAttr = new CellAttr();
                                if (i == e.select("td").size()) {
                                    cellAttr.setEtype(10);
                                }
                                cellAttr.setRowid(detailRow);
                                cellAttr.setColid(detailCol);
                                LinkedHashMap map = new LinkedHashMap();
                                pluginData.put("" + detailCol, map);
                                parseSheet.buildPluginCellMap(cellAttr, map);
                                pluginDetailtableData.put(detailRow + "", pluginData);
                                detailCol++;
                            }
                            detailRowheads.put("row_" + detailRow, "30");
                            pluginData = new LinkedHashMap();
                            detailRow++;
                            detailCol = 0;
                        }
                        if (detailCol == 0) {
                            CellAttr cellAttr = new CellAttr();
                            cellAttr.setRowid(detailRow);
                            cellAttr.setColid(detailCol);
                            cellAttr.setEvalue("全选");
                            cellAttr.setEtype(20);
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
                            detail.add(map);
                            detailMap.put("ec", detail);
                            detailColheads.put("col_" + detailCol, "50");
                            parseSheet.buildDataEcMap(cellAttr, map);
                            parseSheet.buildPluginCellMap(cellAttr, map2);
                            pluginData.put("" + detailCol, map2);
                            detailCol++;
                        }
                        CellAttr cellAttr = new CellAttr();
                        cellAttr.setEvalue(detailtd.text());
                        cellAttr.setRowid(detailRow);
                        cellAttr.setColid(detailCol);
                        cellAttr.setColspan(1);
                        cellAttr.setRowspan(1);
                        cellAttr.setEtype(1); // 1,文本；2,字段名；3,表单内容
                        cellAttr.setFieldattr(1); // 编辑
                        cellAttr.setEvalue(detailtd.text());
                        cellAttr.setHalign(1); // 左右居中
                        cellAttr.setValign(1); // 上下居中
                        cellAttr.setBackground_color("#e7f3fc");
                        cellAttr.setBtop_style(1);
                        cellAttr.setBtop_color("#90badd");
                        cellAttr.setBbottom_style(1);
                        cellAttr.setBbottom_color("#90badd");
                        cellAttr.setBleft_style(1);
                        cellAttr.setBleft_color("#90badd");
                        cellAttr.setBright_style(1);
                        cellAttr.setBright_color("#90badd");

                        if (cellAttr.getEvalue().contains("序号")) {
                            detailColheads.put("col_" + detailCol, "50");
                        } else {
                            detailColheads.put("col_" + detailCol, "120");
                        }
                        LinkedHashMap map = new LinkedHashMap();
                        LinkedHashMap map2 = new LinkedHashMap();
                        parseSheet.buildDataEcMap(cellAttr, map);
                        parseSheet.buildPluginCellMap(cellAttr, map2);
                        pluginData.put("" + detailCol, map2);
                        detail.add(map);
                        detailMap.put("ec", detail);
                        detailCol++;
                    }
                    pluginDetailtableData.put(detailRow + "", pluginData);

                    if (detailtd.nextElementSibling() == null) {
                        maxColumn = detailCol;
                        detailRowheads.put("row_" + detailRow, "30");
                        detailRow++; // 下一行
                        detailCol = 0; // 列归零
                        detailMap.put("rowheads", detailRowheads);
                    }
                }
                pluginData = new LinkedHashMap();
                CellAttr cellAttr = new CellAttr();
                LinkedHashMap map = new LinkedHashMap();
                LinkedHashMap map2 = new LinkedHashMap();
                if (((ArrayList) detailMap.get("ec")).size() > maxColumn * 2) {
                    cellAttr.setEtype(9);
                    cellAttr.setEvalue("表尾标识");
                    cellAttr.setBackground_color("#eeeeee");
                    detailMap.put("edtailinrow", detailRow);
                    LinkedHashMap mergedCell = new LinkedHashMap();
                    mergedCell.put("row", detailRow);
                    mergedCell.put("rowCount", 1);
                    mergedCell.put("col", detailCol);
                    mergedCell.put("colCount", maxColumn);
                    detailSpans.add(mergedCell);
                } else {
                    cellAttr.setEtype(8);
                    cellAttr.setEvalue("表头标识");
                    cellAttr.setBackground_color("#eeeeee");
                    detailMap.put("edtitleinrow", detailRow);
                    LinkedHashMap mergedCell = new LinkedHashMap();
                    mergedCell.put("row", detailRow);
                    mergedCell.put("rowCount", 1);
                    mergedCell.put("col", detailCol);
                    mergedCell.put("colCount", maxColumn);
                    detailSpans.add(mergedCell);
                }
                pluginData.put(detailCol, map2);
                cellAttr.setRowid(detailRow);
                cellAttr.setColid(detailCol);
                parseSheet.buildDataEcMap(cellAttr, map);
                parseSheet.buildPluginCellMap(cellAttr, map2);
                detail.add(map);
                detailMap.put("ec", detail);
                detailMap.put("seniorset", "1");
                detailMap.put("colheads", detailColheads);
                detailRowheads.put("row_" + detailRow, "30");
                pluginDetailtableData.put(detailRow + "", pluginData);
                detailRow++;
            }
            detailRow = 0;
            detailColheads = new LinkedHashMap<>();
            return detailMap;
        }
        for (Element e : element.children()) {
            detailTableAnalyse(e);
        }
        return detailMap;
    }


//    /**
//     * 取出表单关联的js,css
//     *
//     * @param element
//     * @return
//     */
//    int count = 0;
//
//    public LinkedHashMap getCssAndScript(Element element) {
//        LinkedHashMap linkMap = new LinkedHashMap();
//        LinkedHashMap scriptMap = new LinkedHashMap();
//        int i = 0;
//
//        if (element.select("link").size() != 0
//                || element.select("script").size() != 0) {
//            for (Element e : element.select("link")) {
//                linkMap.put(i++, e.toString());
//            }
//            i = 0;
//            for (Element e : element.select("script")) {
//                scriptMap.put(i++, e.toString());
//            }
//        }
//        formula.put("link", linkMap);
//        formula.put("script", scriptMap);
//        return formula;
//    }


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


    /**
     * 构造明细表的pluginjson
     */
    private void getPluginMap(String str, LinkedHashMap pluginMap, LinkedHashMap data) {
        pluginMap.put("version", "2.0");
        pluginMap.put("tabStripVisible", false);
        pluginMap.put("canUserEditFormula", false);
        pluginMap.put("allowUndo", false);
        pluginMap.put("allowDragDrop", false);
        pluginMap.put("allowDragFill", false);
        pluginMap.put("grayAreaBackColor", "white");
        LinkedHashMap<String, Object> sheets = new LinkedHashMap<String, Object>();
        pluginMap.put("sheets", sheets);
        LinkedHashMap<String, Object> Sheet1 = new LinkedHashMap<String, Object>();
        sheets.put("Sheet1", Sheet1);

        Sheet1.put("name", "Sheet1");
        LinkedHashMap<String, Object> defaults = new LinkedHashMap<String, Object>();
        defaults.put("rowHeight", 30);
        defaults.put("colWidth", 62);
        defaults.put("rowHeaderColWidth", 40);
        defaults.put("colHeaderRowHeight", 20);
        Sheet1.put("defaults", defaults);
        LinkedHashMap<String, Object> selections = new LinkedHashMap<String, Object>();
        LinkedHashMap<String, Object> selections_0 = new LinkedHashMap<String, Object>();
        selections.put("0", selections_0);
        selections_0.put("row", 0);
        selections_0.put("rowCount", 1);
        selections_0.put("col", 0);
        selections_0.put("colCount", 1);
        Sheet1.put("selections", selections);
        Sheet1.put("activeRow", 0);
        Sheet1.put("activeCol", 0);
        if (str.contains("main")) {
            Sheet1.put("spans", mainSpans);
        } else if (str.contains("detail")) {
            Sheet1.put("spans", detailSpans);
        }
        LinkedHashMap<String, Object> gridline = new LinkedHashMap<String, Object>();
        gridline.put("color", "#D0D7E5");
        gridline.put("showVerticalGridline", true);
        gridline.put("showHorizontalGridline", true);
        Sheet1.put("gridline", gridline);
        Sheet1.put("allowDragDrop", false);
        Sheet1.put("allowDragFill", false);

        LinkedHashMap<String, Object> rowHeaderData = new LinkedHashMap<String, Object>();
        LinkedHashMap<String, Object> colHeaderData = new LinkedHashMap<String, Object>();
        LinkedHashMap<String, Object> defaultDataNode = new LinkedHashMap<String, Object>();
        LinkedHashMap<String, Object> defaultDataNode_style = new LinkedHashMap<String, Object>();
        defaultDataNode_style.put("foreColor", "black");
        defaultDataNode.put("style", defaultDataNode_style);
        rowHeaderData.put("rowCount", mainRow > 20 ? mainRow + 5 : 20);
        rowHeaderData.put("defaultDataNode", defaultDataNode);
        colHeaderData.put("colCount", mainCol > 10 ? mainCol + 3 : 10);
        colHeaderData.put("defaultDataNode", defaultDataNode);
        Sheet1.put("rowHeaderData", rowHeaderData);
        Sheet1.put("colHeaderData", colHeaderData);

        LinkedHashMap dataTable = new LinkedHashMap();
        dataTable.put("dataTable", data);
        Sheet1.put("columns", getPluginColumns(str, data));
//        Sheet1.put("mainSpans", getPlugin_combine());
        Sheet1.put("data", dataTable);


        LinkedHashMap<String, Object> rowRangeGroup = new LinkedHashMap<String, Object>();
        rowRangeGroup.put("itemsCount", mainRow > 20 ? mainRow : 20);
        LinkedHashMap<String, Object> colRangeGroup = new LinkedHashMap<String, Object>();
        colRangeGroup.put("itemsCount", mainCol > 10 ? mainCol : 10);
        Sheet1.put("rowRangeGroup", rowRangeGroup);
        Sheet1.put("colRangeGroup", colRangeGroup);

        plugin.put(str, pluginMap);
    }


    /**
     * 为明细表的每一列配置信息
     */
    private ArrayList getPluginColumns(String str, LinkedHashMap map) {
        ArrayList list = new ArrayList();
        if (str.contains("main")) {
            for (int i = 0; i < totalCols; i++) {
                LinkedHashMap<String, Object> plugin_column = new LinkedHashMap();
                plugin_column.put("size", 120);
                plugin_column.put("dirty", "true");
                list.add(plugin_column);
            }
        } else {
            for (int i = 0; i < ((LinkedHashMap) map.get(3 + "")).size(); i++) {
                LinkedHashMap<String, Object> plugin_column = new LinkedHashMap();
                if (i == 0 || i == 1) { // 单选框和序号列宽度设为50
                    plugin_column.put("size", 50);
                    plugin_column.put("dirty", "true");
                } else {
                    plugin_column.put("size", 120);
                    plugin_column.put("dirty", "true");
                }
                list.add(plugin_column);
            }
        }
        return list;
    }


//    private ArrayList getPlugin_combine() {
//        ArrayList list = new ArrayList();
//        LinkedHashMap<String, Object> plugin_column = new LinkedHashMap<String, Object>();
//        plugin_column.put("row", "0");
//        plugin_column.put("rowCount", "1");
//        plugin_column.put("col", "0");
//        plugin_column.put("colCount", "6");
//        list.add(plugin_column);
//        return list;
//    }

    /**
     * 在主表中生成明细索引
     */
    private void getdetailTableCell(String str) {
        CellAttr cellAttr = new CellAttr();
        cellAttr.setRowid(++mainRow);
        cellAttr.setColid(mainCol);
        cellAttr.setColspan(totalCols);
        cellAttr.setEtype(7); // 明细
        cellAttr.setEvalue(str.substring(7));
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
        parseSheet.buildDataEcMap(cellAttr, map);
        mainList.add(map);
        emaintable.put("ec", mainList);
        mainRowheads.put("row_" + mainRow, "30");
    }


    /**
     * 获取合并单元格信息
     */
    private void getMergedCellsInfo(int row, int col, int span, String flag) {
        if (flag.equals("colspan")) {
            LinkedHashMap mergedCell = new LinkedHashMap();
            mergedCell.put("row", row);
            mergedCell.put("rowCount", 1);
            mergedCell.put("col", col);
            mergedCell.put("colCount", span);
            mainSpans.add(mergedCell);
        } else if (flag.equals("rowspan")) {
            LinkedHashMap mergedCell = new LinkedHashMap();
            mergedCell.put("row", row);
            mergedCell.put("rowCount", span);
            mergedCell.put("col", col);
            mergedCell.put("colCount", 1);
            mainSpans.add(mergedCell);
        }
    }


    /**
     * 获取主表colspan和rowspan信息
     */
    private void getMainTableSpans(CellAttr cellAttr, Element e) {
        boolean hasColspan = e.attr("colspan").equals("");
        boolean hasRowspan = e.attr("rowspan").equals("");
        if (!hasColspan) {
            getMergedCellsInfo(mainRow, mainCol, Integer.valueOf(e.attr("colspan")), "colspan");
            cellAttr.setColspan(Integer.valueOf(e.attr("colspan")));
        } else {
            cellAttr.setColspan(1);
        }
        if (!hasRowspan) {
            getMergedCellsInfo(mainRow, mainCol, Integer.valueOf(e.attr("colspan")), "rowspan");
            cellAttr.setRowspan(Integer.valueOf(e.attr("rowspan")));
        } else {
            cellAttr.setRowspan(1);
        }
    }
}

