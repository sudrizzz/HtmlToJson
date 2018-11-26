import net.sf.json.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map;

public class CellAnalysis {

    int mainRow = 0; // 主表行
    int mainCol = 0; // 主表列
    int detailRow = 0; // 明细表行
    int detailCol = 0; // 明细表列
    int detailCount = 0; // 明细计数
    int totalCols = 6; // 主表单元格数
    ArrayList mainList = new ArrayList(); // 主表list
    ArrayList detail; // 明细表list
    ArrayList mainSpans = new ArrayList(); // 主表合并单元格信息
    ArrayList detailSpans = new ArrayList(); // 明细合并单元格信息
    StringBuffer stringBuffer = new StringBuffer();
    LinkedHashMap eattr = new LinkedHashMap<>(); // 流程信息
    //    LinkedHashMap formula = new LinkedHashMap(); // 脚本和公式
    LinkedHashMap emaintable = new LinkedHashMap(); // 主表字段
    LinkedHashMap etables = new LinkedHashMap(); // 表单内容
    LinkedHashMap rowattrs = new LinkedHashMap(); // 行隐藏
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
    ExcelSecurity excelSecurity = new ExcelSecurity();

    public LinkedHashMap getTableInfo(StringBuffer sb) {
        Document document = Jsoup.parse(sb.toString());
        int totalWidth = 1024;
        if (document.hasClass("table") && document.selectFirst("table").hasAttr("width")) {
            if (!"".equals(document.selectFirst("table[width]").attr("width"))) {
                totalWidth = Integer.valueOf(document.selectFirst("table[width]").attr("width"));
            }
        }
        if (document.hasClass("td[colspan]") && document.selectFirst("td[colspan]").hasAttr("colspan")) {
            if (!"".equals(document.selectFirst("td[colspan]").attr("colspan"))) {
                totalCols = Integer.valueOf(document.selectFirst("td[colspan]").attr("colspan"));
            }
        }
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
        jsonMap.put("datajson", JSONObject.fromObject(datajsonMap));
        jsonMap.put("pluginjson", JSONObject.fromObject(plugin));
        jsonMap.put("script", excelSecurity.encodeStr(stringBuffer.toString()));
        return jsonMap;
    }


    /**
     * 递归取出所有的节点信息，包括字段显示名和标题
     */
    public LinkedHashMap recursiveDFS(Element element) {

        if (element.tagName().equals("tr") && element.text().trim().equals("")) {
            return dataJsonMap;
        }
        if (element.tagName().equals("script")) {
            stringBuffer.append(element.toString());
        }
        if (element.tagName().equals("tr")
                && element.select("table").size() == 0
                && element.child(0).tagName().equals("td")) {
            if (((ArrayList) emaintable.get("ec")).size() > 0) {
                mainRow++; // 行数加一
            }
            mainCol = 0; // 列归零
            mainRowheads.put("row_" + mainRow, "30");
            if (!mainRowheads.containsKey("row_0")) {
                mainRowheads.put("row_0", "60");
            }
        }
        if (element.tagName().equals("tr")) {
            if (checkForDisplayNone(element)) {
                LinkedHashMap hideattr = new LinkedHashMap();
                hideattr.put("hide", "y");
                rowattrs.put("row_" + mainRow, hideattr);
                emaintable.put("rowattrs", rowattrs);
            }
            LinkedHashMap pluginData = new LinkedHashMap();
            for (int i = 0; i < element.select("td").size(); i++) {
                Element e = element.select("td").get(i);
                if (e.tagName().equals("td")) {
                    if (e.select("table").size() > 0) { // 如果td里面是table，则认为其是明细表，进行特殊处理
                        if (e.selectFirst("table").select("strong").size() > 0
                                && e.selectFirst("table").select("button").size() > 0) {
                            Element detailTitle = e.selectFirst("table").selectFirst("strong").parent();
                            if (!detailTitle.text().equalsIgnoreCase("&nbsp;")
                                    && !detailTitle.text().equalsIgnoreCase("&nbsp")
                                    && !detailTitle.text().equalsIgnoreCase("")) {
                                CellAttr cellAttr = createCell();
                                cellAttr.setFont_size("12pt");
                                cellAttr.setBold(true);
                                cellAttr.setRowid(++mainRow);
                                cellAttr.setColid(0);
                                cellAttr.setEtype(1); // 1,文本；2,字段名；3,表单内容
                                cellAttr.setEvalue(detailTitle.text());
                                cellAttr.setFont_color(detailTitle.attr("color"));
                                getMainTableSpans(cellAttr, detailTitle);

                                LinkedHashMap map = new LinkedHashMap();
                                LinkedHashMap map2 = new LinkedHashMap();
                                parseSheet.buildDataEcMap(cellAttr, map);
                                parseSheet.buildPluginCellMap(cellAttr, map2);
                                pluginData.put("" + 0, map2);
                                pluginMaintableData.put(mainRow + "", pluginData);
                                mainList.add(map);
                                emaintable.put("ec", mainList);
                                mainRowheads.put("row_" + mainRow, "30");
                            }
                        }
                        for (Element element1 : element.select("table")) {
                            // 如果该元素下面还有table元素，则继续循环，直到取到最底层的明细table，防止重复
                            if (element1.select("table").size() > 1) {
                                continue;
                            }
                            if (element1.select("strong").size() == 0 && element1.select("input").size() == 0) {
                                continue;
                            }
                            detail = new ArrayList();
                            LinkedHashMap map = detailTableAnalysis(element1);
                            if (map.size() > 0) {
                                etables.put("detail_" + ++detailCount, map);
                                getPluginMap("detail_" + detailCount + "_sheet", new LinkedHashMap(), pluginDetailtableData);
                                getdetailTableCell("detail_" + detailCount);

                                CellAttr cellAttr = createCell();
                                cellAttr.setRowid(mainRow);
                                cellAttr.setColid(0);
                                cellAttr.setEvalue("明细表" + detailCount);
                                cellAttr.setBackground_color("#e7f3fc");
                                cellAttr.setHalign(0);
                                getMainTableSpans(cellAttr, e);

                                LinkedHashMap map2 = new LinkedHashMap();
                                pluginData = new LinkedHashMap();
                                parseSheet.buildPluginCellMap(cellAttr, map2);
                                pluginData.put("" + 0, map2);
                                pluginMaintableData.put(mainRow + "", pluginData);
                                detailSpans = new ArrayList(); // 重置明细表的列合并信息
                            }
                        }
                        return etables;
                        // 如果td里面是input框，则针对input框进行处理
                    } else if (e.select("input").size() > 0) {
                        for (int j = 0; j < e.select("input").size(); j++) {
                            CellAttr cellAttr = createCell();
                            String attr = e.select("input").get(j).val();
                            if (attr.contains("序号")) {
                                cellAttr.setEtype(22); // 22,序号
                            } else {
                                cellAttr.setEtype(3); // 2,字段名；3,表单内容
                            }
                            if (attr.contains("[") && attr.contains("]")) {
                                cellAttr.setEvalue(attr.substring(attr.indexOf("]") + 1));
                            } else {
                                cellAttr.setEvalue(attr);
                            }
                            cellAttr.setRowid(mainRow);
                            cellAttr.setColid(mainCol);
                            cellAttr.setFieldid(e.select("input").get(j).attr("name").substring(5));
                            cellAttr.setFieldtype(e.select("input").get(j).attr("type")); // 单行文本框
                            getMainTableSpans(cellAttr, e);

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
                        CellAttr cellAttr = createCell();
                        if (e.select("font").size() > 0) {
                            cellAttr.setFont_size(e.selectFirst("font").attr("size"));
                        } else {
                            cellAttr.setFont_size("12pt");
                        }
                        cellAttr.setBold(true);
                        cellAttr.setRowid(mainRow);
                        cellAttr.setColid(0);
                        cellAttr.setEtype(1); // 1,文本；2,字段名；3,表单内容
                        cellAttr.setEvalue(e.text());
                        if (e.select("font").size() > 0) {
                            cellAttr.setFont_color(e.selectFirst("font").attr("color"));
                        } else {
                            cellAttr.setFont_color(e.attr("color"));
                        }
                        getMainTableSpans(cellAttr, e);

                        LinkedHashMap map = new LinkedHashMap();
                        LinkedHashMap map2 = new LinkedHashMap();
                        parseSheet.buildDataEcMap(cellAttr, map);
                        parseSheet.buildPluginCellMap(cellAttr, map2);
                        pluginData.put("" + mainCol, map2);
                        mainList.add(map);
                        emaintable.put("ec", mainList);
                    } else if (e.select("a").size() > 0) { // 超链接
                        CellAttr cellAttr = createCell();
                        cellAttr.setRowid(mainRow);
                        cellAttr.setColid(mainCol);
                        cellAttr.setEtype(11); // 11,超链接
                        cellAttr.setEvalue(e.text());
                        cellAttr.setFont_color("#0000ee");
                        getMainTableSpans(cellAttr, e);

                        LinkedHashMap map = new LinkedHashMap();
                        LinkedHashMap map2 = new LinkedHashMap();
                        parseSheet.buildDataEcMap(cellAttr, map);
                        parseSheet.buildPluginCellMap(cellAttr, map2);
                        pluginData.put("" + mainCol, map2);
                        mainList.add(map);
                        emaintable.put("ec", mainList);
                    } else { // 否则就为普通的td
                        CellAttr cellAttr = createCell();
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
                        cellAttr.setBackground_color("#e7f3fc");
                        getMainTableSpans(cellAttr, e);

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
    private LinkedHashMap detailTableAnalysis(Element element) {
        LinkedHashMap detailMap = new LinkedHashMap();
        if (element.select("strong").size() > 0) {
            if (element.select("input").size() <= 1 && mainList.size() == 0) {
                LinkedHashMap pluginData = new LinkedHashMap();
                Element e = element.selectFirst("strong");
                if ((e.text().trim().equals("")) || (e.text().trim().equals("&nbsp;"))) {
                    return etables;
                }
                if (e.tagName().equals("strong")) {
                    e = e.parent();
                    Element colspanElement = e;
                    CellAttr cellAttr = createCell();
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
                    cellAttr.setFont_size("24pt");
                    cellAttr.setBold(true);
                    cellAttr.setRowspan(e.attr("rowspan").equals("") ? 1 : Integer.valueOf(e.attr("rowspan")));
                    cellAttr.setEtype(1); // 2,字段名；3,表单内容
                    cellAttr.setEvalue(e.text());
                    cellAttr.setFont_color(e.attr("color"));

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
                if (e.childNodeSize() <= 1) {
                    continue;
                }
                for (Element detailtd : e.select("td")) {
                    if (checkForDisplayNone(detailtd)) {
                        continue;
                    }
                    if (detailtd.select("input").size() > 0) {
                        if (detailCol == 0) {
                            CellAttr cellAttr = createCell();
                            cellAttr.setRowid(detailRow);
                            cellAttr.setColid(detailCol);
                            cellAttr.setEtype(21);
                            cellAttr.setEvalue("选中");

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
                            CellAttr cellAttr = createCell();
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
                            if (attr.contains("[") && attr.contains("]")) {
                                cellAttr.setEvalue(attr.substring(attr.indexOf("]") + 1));
                            } else {
                                cellAttr.setEvalue(attr);
                            }

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
                            int count = e.select("td").size() - e.select("td[style=display: none]").size() + 1;
                            for (int i = 0; i < count; i++) {
                                CellAttr cellAttr = new CellAttr();
                                // 明细空一行的最后一格显示添加和删除按钮
                                if (i == count - 1) {
                                    cellAttr.setEtype(10);
                                    cellAttr.setBright_style(1);
                                    cellAttr.setBright_color("#90badd");
                                }
                                cellAttr.setRowid(detailRow);
                                cellAttr.setColid(detailCol);
                                LinkedHashMap map = new LinkedHashMap();
                                LinkedHashMap map2 = new LinkedHashMap();
                                detail.add(map);
                                detailMap.put("ec", detail);
                                pluginData.put("" + detailCol, map);
                                parseSheet.buildDataEcMap(cellAttr, map);
                                parseSheet.buildPluginCellMap(cellAttr, map2);
                                pluginDetailtableData.put(detailRow + "", pluginData);
                                detailCol++;
                            }
                            detailRowheads.put("row_" + detailRow, "30");
                            pluginData = new LinkedHashMap();
                            detailRow++;
                            detailCol = 0;
                        }
                        if (detailCol == 0) {
                            CellAttr cellAttr = createCell();
                            cellAttr.setRowid(detailRow);
                            cellAttr.setColid(detailCol);
                            cellAttr.setEvalue("全选");
                            cellAttr.setEtype(20);
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
                        CellAttr cellAttr = createCell();
                        cellAttr.setRowid(detailRow);
                        cellAttr.setColid(detailCol);
                        cellAttr.setColspan(1);
                        cellAttr.setRowspan(1);
                        cellAttr.setEtype(2); // 1,文本；2,字段名；3,表单内容
                        cellAttr.setEvalue(detailtd.text());
                        Element lastTr = detailtd.parent().lastElementSibling();
                        while (lastTr != null && lastTr.select("td[style=display: none]").size() > 0
                                && lastTr.select("td[style=display: none]").size() == lastTr.select("td").size()) {
                            lastTr = lastTr.previousElementSibling();
                        }
                        if (lastTr != null && lastTr.select("td").get(detailCol - 1).select("input").size() > 0) {
                            cellAttr.setFieldid(lastTr.select("td").get(detailCol - 1).selectFirst("input").attr("name").substring(5));
                        }
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
                    int index = (detailtd.siblingIndex() + 1) / 2 - 1;
                    pluginDetailtableData.put(detailRow + "", pluginData);
                    if (detailtd.nextElementSibling() == null
                            || (checkForDisplayNone(detailtd.nextElementSibling())
                            && e.select("td:gt(" + index + ")").size() == e.select("td[style=display: none]:gt(" + index + ")").size())) {
                        maxColumn = detailCol;
                        detailRowheads.put("row_" + detailRow, "30");
                        detailRow++; // 下一行
                        detailCol = 0; // 列归零
                        detailMap.put("rowheads", detailRowheads);
                    }
                }
                pluginData = new LinkedHashMap();
                CellAttr cellAttr = createCell();
                LinkedHashMap map = new LinkedHashMap();
                LinkedHashMap map2 = new LinkedHashMap();
                ArrayList detailMapList = (ArrayList) detailMap.get("ec");
                if (!((Map) detailMapList.get(detailMapList.size() - 1)).get("evalue").equals("表尾标识")) {
                    if (detailMapList.size() > maxColumn * 2) {
                        cellAttr.setEtype(9);
                        cellAttr.setEvalue("表尾标识");
                        cellAttr.setValign(1);
                        cellAttr.setBackground_color("#eeeeee");
                        detailMap.put("edtailinrow", detailRow + "");
                        LinkedHashMap mergedCell = new LinkedHashMap();
                        mergedCell.put("row", detailRow);
                        mergedCell.put("rowCount", 1);
                        mergedCell.put("col", detailCol);
                        mergedCell.put("colCount", maxColumn);
                        detailSpans.add(mergedCell);
                    } else {
                        cellAttr.setEtype(8);
                        cellAttr.setEvalue("表头标识");
                        cellAttr.setValign(1);
                        cellAttr.setBackground_color("#eeeeee");
                        detailMap.put("edtitleinrow", detailRow + "");
                        LinkedHashMap mergedCell = new LinkedHashMap();
                        mergedCell.put("row", detailRow);
                        mergedCell.put("rowCount", 1);
                        mergedCell.put("col", detailCol);
                        mergedCell.put("colCount", maxColumn);
                        detailSpans.add(mergedCell);
                    }
                }
                if (cellAttr.getEtype() != 0) {
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
            }
            detailRow = 0;
            detailColheads = new LinkedHashMap<>();
            return detailMap;
        }
        for (Element e : element.children()) {
            detailTableAnalysis(e);
        }
        return detailMap;
    }


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
        defaults.put("colWidth", 60);
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
        LinkedHashMap<String, Object> gridline = new LinkedHashMap<String, Object>();
        gridline.put("color", "#D0D7E5");
        gridline.put("showVerticalGridline", true);
        gridline.put("showHorizontalGridline", true);
        Sheet1.put("gridline", gridline);
        Sheet1.put("allowDragDrop", false);
        Sheet1.put("allowDragFill", false);

        LinkedHashMap<String, Object> dataTable = new LinkedHashMap();
        LinkedHashMap<String, Object> rowHeaderData = new LinkedHashMap<String, Object>();
        LinkedHashMap<String, Object> colHeaderData = new LinkedHashMap<String, Object>();
        LinkedHashMap<String, Object> defaultDataNode = new LinkedHashMap<String, Object>();
        LinkedHashMap<String, Object> rowRangeGroup = new LinkedHashMap<String, Object>();
        LinkedHashMap<String, Object> colRangeGroup = new LinkedHashMap<String, Object>();
        LinkedHashMap<String, Object> defaultDataNode_style = new LinkedHashMap<String, Object>();

        defaultDataNode_style.put("foreColor", "black");
        defaultDataNode.put("style", defaultDataNode_style);

        Sheet1.put("rowHeaderData", rowHeaderData);
        Sheet1.put("colHeaderData", colHeaderData);
        if (str.contains("main")) {
            rowHeaderData.put("rowCount", mainRow > 20 ? mainRow + 5 : 25);
            colHeaderData.put("colCount", mainCol > 10 ? mainCol + 3 : 15);
            rowHeaderData.put("defaultDataNode", defaultDataNode);
            colHeaderData.put("defaultDataNode", defaultDataNode);
            rowRangeGroup.put("itemsCount", mainRow > 20 ? mainRow + 5 : 25);
            colRangeGroup.put("itemsCount", mainCol > 10 ? mainCol + 3 : 15);
            Sheet1.put("rowRangeGroup", rowRangeGroup);
            Sheet1.put("colRangeGroup", colRangeGroup);
            Sheet1.put("spans", mainSpans);
            dataTable.put("rowCount", mainRow > 20 ? mainRow + 5 : 25);
            dataTable.put("colCount", mainCol > 10 ? mainCol + 3 : 10);
        } else if (str.contains("detail")) {
            int rowCount = data.size();
            int colCount = ((LinkedHashMap)data.get("3")).size();
            rowHeaderData.put("rowCount", rowCount > 5 ? rowCount + 3 : 8);
            colHeaderData.put("colCount", colCount > 20 ? colCount + 5 : 25);
            rowHeaderData.put("defaultDataNode", defaultDataNode);
            colHeaderData.put("defaultDataNode", defaultDataNode);
            rowRangeGroup.put("itemsCount", rowCount > 5 ? rowCount + 3 : 8);
            colRangeGroup.put("itemsCount", colCount > 20 ? colCount + 5 : 25);
            Sheet1.put("rowRangeGroup", rowRangeGroup);
            Sheet1.put("colRangeGroup", colRangeGroup);
            Sheet1.put("spans", detailSpans);
            dataTable.put("rowCount", rowCount > 5 ? rowCount + 3 : 8);
            dataTable.put("colCount", colCount > 20 ? colCount + 5 : 25);
        }
        dataTable.put("dataTable", data);
        Sheet1.put("columns", getPluginColumns(str, data));
        Sheet1.put("data", dataTable);
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
                plugin_column.put("dirty", true);
                list.add(plugin_column);
            }
        } else {
            boolean isSerialNum = ((LinkedHashMap)((LinkedHashMap) map.get("3")).get("1")).containsValue("序号");
            for (int i = 0; i < ((LinkedHashMap) map.get(3 + "")).size(); i++) {
                LinkedHashMap<String, Object> plugin_column = new LinkedHashMap();
                if (i == 0 || (i == 1 && isSerialNum)) { // 单选框和序号列宽度设为50
                    plugin_column.put("size", 50);
                    plugin_column.put("dirty", true);
                } else {
                    plugin_column.put("size", 120);
                    plugin_column.put("dirty", true);
                }
                list.add(plugin_column);
            }
        }
        return list;
    }


    /**
     * 在主表中生成明细索引
     */
    private void getdetailTableCell(String str) {
        CellAttr cellAttr = createCell();
        cellAttr.setRowid(++mainRow);
        cellAttr.setColid(0);
        cellAttr.setColspan(totalCols);
        cellAttr.setEtype(7); // 明细
        cellAttr.setEvalue(str.substring(7));

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
        if (e.childNodeSize() == 1 && e.select("strong").size() > 0) {
            cellAttr.setColspan(totalCols);
            cellAttr.setRowspan(1);
            getMergedCellsInfo(mainRow, 0, totalCols, "colspan");
            return;
        }
        boolean hasColspan = !e.attr("colspan").equals("");
        boolean hasRowspan = !e.attr("rowspan").equals("");
        if (hasColspan) {
            if (mainCol == totalCols) {
                mainCol = 0;
            }
            if (mainCol + Integer.valueOf(e.attr("colspan")) > totalCols) {
                getMergedCellsInfo(mainRow, mainCol, totalCols - mainCol, "colspan");
                cellAttr.setColspan(totalCols - mainCol);
            } else {
                getMergedCellsInfo(mainRow, mainCol, Integer.valueOf(e.attr("colspan")), "colspan");
                cellAttr.setColspan(Integer.valueOf(e.attr("colspan")));
            }
        } else {
            cellAttr.setColspan(1);
        }
        if (hasRowspan) {
            if (mainCol == totalCols) {
                mainCol = 0;
            }
            getMergedCellsInfo(mainRow, mainCol, Integer.valueOf(e.attr("rowspan")), "rowspan");
            cellAttr.setRowspan(Integer.valueOf(e.attr("rowspan")));
        } else {
            cellAttr.setRowspan(1);
        }
    }

    /**
     * 检查style=display:none的元素
     */
    private boolean checkForDisplayNone(Element element){
        if (element.hasAttr("style")
                && (element.attr("style").equals("display: none")
                || element.attr("style").equals("display:none"))) {
            return true;
        }
        return false;
    }


    /**
     * 构建单元格并配置属性
     */
    private CellAttr createCell() {
        CellAttr cellAttr = new CellAttr();
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
        return cellAttr;
    }
}

