import net.sf.json.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.util.ArrayList;
import java.util.LinkedHashMap;

public class CellAnalysis {

    int mainRow = 0; // 主表行
    int mainCol = 0; // 主表列
    int detailRow = 0; // 明细表行
    int detailCol = 0; // 明细表列
    int detailCount = 0; // 明细计数
    int totalCols = 6; // 主表一行单元格数，默认为 6
    int[][] mainSpanArray = new int[200][50]; // 用于记录主表字段的行合并
    ArrayList<LinkedHashMap> mainList = new ArrayList(); // 主表list
    ArrayList<LinkedHashMap> detailList = new ArrayList(); // 明细表list
    ArrayList<LinkedHashMap> mainSpans = new ArrayList(); // 主表合并单元格信息
    ArrayList<LinkedHashMap> detailSpans = new ArrayList(); // 明细合并单元格信息
    StringBuffer scripts = new StringBuffer(); // 加密后的 javascript
    LinkedHashMap eattr = new LinkedHashMap<>(); // 流程信息
    LinkedHashMap emaintable = new LinkedHashMap(); // 主表字段
    LinkedHashMap etables = new LinkedHashMap(); // 表单内容
    LinkedHashMap rowattrs = new LinkedHashMap(); // 行隐藏
    LinkedHashMap<String, Object> datajson = new LinkedHashMap(); // datajson
    LinkedHashMap<String, Object> pluginjson = new LinkedHashMap(); // pluginjson
    LinkedHashMap<String, Object> eformdesign = new LinkedHashMap(); // 前台显示的表单信息
    LinkedHashMap<String, Object> mainRowheads = new LinkedHashMap(); // 主表行宽
    LinkedHashMap<String, Object> mainColheads = new LinkedHashMap(); // 主表列宽
    LinkedHashMap<String, Object> detailRowheads = new LinkedHashMap(); // 明细表行高
    LinkedHashMap<String, Object> detailColheads = new LinkedHashMap(); // 明细表列宽
    LinkedHashMap<String, Object> pluginMaintableData = new LinkedHashMap(); // pluginjson 主表
    LinkedHashMap<String, Object> pluginDetailtableData = new LinkedHashMap(); // pluginjson 明细表

    ParseSheet parseSheet = new ParseSheet();
    ExcelSecurity excelSecurity = new ExcelSecurity();

    public LinkedHashMap getTableInfo(StringBuffer sb) {
        Document document = Jsoup.parse(sb.toString());
        int totalWidth = 1000;
        if (document.hasClass("table") && document.selectFirst("table").hasAttr("width")) {
            if (!"".equals(document.selectFirst("table[width]").attr("width"))) {
                totalWidth = Integer.valueOf(document.selectFirst("table[width]").attr("width"));
            }
        }
        if (document.select("td[colspan]").size() > 0 && document.selectFirst("td[colspan]").hasAttr("colspan")) {
            String colCountStr = document.selectFirst("td[colspan]").attr("colspan");
            if (!"".equals(colCountStr)) {
                totalCols = Integer.valueOf(colCountStr) > totalCols ? Integer.valueOf(colCountStr) : totalCols;
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
        datajson.put("eformdesign", eformdesign);
        LinkedHashMap datajsonMap = recursiveDFS(document);
        getPluginMap("main_sheet", new LinkedHashMap(), pluginMaintableData);

        LinkedHashMap jsonMap = new LinkedHashMap();
        jsonMap.put("datajson", JSONObject.fromObject(datajsonMap));
        jsonMap.put("pluginjson", JSONObject.fromObject(pluginjson));
        jsonMap.put("script", excelSecurity.encodeStr(scripts.toString()));
        return jsonMap;
    }


    /**
     * 递归取出所有的字段信息
     * 解析的模板文件大致结构为
     * table 主 table
     *     |- tr 主字段
     *     |- tr
     *     |- tr
     *         |- table 次 table
     *             |- tr 次字段
     *             |- tr
     *             |- ...
     */
    private LinkedHashMap recursiveDFS(Element element) {

        if (element.tagName().equals("tr") && element.childNodeSize() == 0) return datajson;
        if (element.tagName().equals("script")) scripts.append(element.toString());
        if (element.tagName().equals("tr")
                && (element.children().size() > 0 && element.child(0).tagName().equals("td"))
                && (element.children().select("tr").size() <= 1 || (element.children().select("table").size() == 1 && element.children().select("table[name*=oTable]").size() == 1))
                && element.children().select("button").size() == 0
        ) {
            if (((ArrayList) emaintable.get("ec")).size() > 0) mainRow++; // 行数加一
            mainCol = 0; // 列归零
            if (mainRow > 1 && !mainRowheads.containsKey("row_" + (mainRow - 1))) {
                mainRowheads.put("row_" + (mainRow - 1), "30");
            }
            mainRowheads.put("row_" + mainRow, "30");
            if (!mainRowheads.containsKey("row_0")) mainRowheads.put("row_0", "60");
        }
        if (element.tagName().equals("tr")) {

            /**
             * 如果当前元素中含有 table ，且 table 的个数与明细 table 的个数不一致
             * 则认为该元素的结构是类似于最顶层元素
             * 所以直接将该元素再次递归
             *
             */
            if (element.children().select("table").size() != element.children().select("table[name*=oTable]").size()) {
                return recursiveDFS(element.child(0));
            }
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
                    if (checkForDisplayNone(e)) {
                        continue;
                    }
                    if (e.select("table[name*=oTable]").size() > 0) {
                        if (e.selectFirst("table").select("strong").size() > 0
                                && e.selectFirst("table").select("button").size() > 0) {
                            Element detailTitle = e.selectFirst("table").selectFirst("strong").parent();
                            if (!detailTitle.text().equalsIgnoreCase("&nbsp;")
                                    && !detailTitle.text().equalsIgnoreCase("&nbsp")
                                    && !detailTitle.text().equalsIgnoreCase("")) {
                                // 明细标题
                                while (mainSpanArray[mainRow + 1][mainCol] == 1) {
                                    mainCol++;
                                }
                                CellAttr cellAttr = createCell();
                                cellAttr.setFont_size("12pt");
                                cellAttr.setBold(true);
                                cellAttr.setRowid(++mainRow);
                                cellAttr.setColid(mainCol);
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
                            if (element1.select("table").size() > 1) continue;
                            if (element1.select("strong").size() == 0 && element1.select("input").size() == 0) continue;

                            detailList = new ArrayList();
                            LinkedHashMap map = detailTableAnalysis(element1);
                            if (map.size() > 0) {
                                etables.put("detail_" + ++detailCount, map);
                                getPluginMap("detail_" + detailCount + "_sheet", new LinkedHashMap(), pluginDetailtableData);
                                getDetailTableCell("detail_" + detailCount);
                                getCoordinate();
                                CellAttr cellAttr = createCell();
                                cellAttr.setRowid(mainRow);
                                cellAttr.setColid(mainCol);
                                cellAttr.setColid(0);
                                cellAttr.setEvalue("明细表" + detailCount);
                                cellAttr.setBackground_color("#e7f3fc");
                                cellAttr.setHalign(0);
                                cellAttr.setBackground_image("/workflow/exceldesign/image/shortBtn/detail/detailTable_wev8.png");
                                cellAttr.setBackground_imagelayout("3");
                                getMainTableSpans(cellAttr, e);

                                LinkedHashMap map2 = new LinkedHashMap();
                                pluginData = new LinkedHashMap();
                                parseSheet.buildPluginCellMap(cellAttr, map2);
                                pluginData.put("" + mainCol, map2);
                                pluginMaintableData.put(mainRow + "", pluginData);
                                detailSpans = new ArrayList(); // 重置明细表的列合并信息
                            }
                        }
                        return etables;
                    } else if (e.getElementsByClass("InputStyle").size() > 0
                            && e.getElementsByTag("td").size() <= 1
                            && mainList.size() > 0) {
                        Elements inputs = e.getElementsByClass("InputStyle");
                        CellAttr cellAttr = createCell();
                        String attr = inputs.get(0).val();
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
                        getCoordinate();
                        cellAttr.setRowid(mainRow);
                        cellAttr.setColid(mainCol);
                        cellAttr.setFieldid(inputs.get(0).attr("name").substring(5));
                        cellAttr.setFieldtype(inputs.get(0).attr("type")); // 单行文本框
                        getMainTableSpans(cellAttr, e);

                        LinkedHashMap map = new LinkedHashMap();
                        LinkedHashMap map2 = new LinkedHashMap();
                        parseSheet.buildDataEcMap(cellAttr, map);
                        parseSheet.buildPluginCellMap(cellAttr, map2);
                        pluginData.put("" + mainCol, map2);
                        mainList.add(map);
                        pluginMaintableData.put(mainRow + "", pluginData);
                        emaintable.put("ec", mainList);
                        if (!e.attr("colspan").equals("")) {
                            mainCol += Integer.valueOf(e.attr("colspan"));
                        } else {
                            mainCol++;
                        }
                        if (mainCol >= totalCols) {
                            mainCol = 0;
                        }
                        // 如果里面有strong标签，则作为标题处理
                    } else if (e.select("strong").size() > 0 && e.select("tr").size() == 0) {
                        if (e.select("input").size() <= 1 && mainList.size() == 0) {
                            if ((e.text().trim().equals("")) || (e.text().trim().equals("&nbsp;"))) {
                                return etables;
                            }
                            e = e.selectFirst("strong");
                            if (e.tagName().equals("strong")) {
                                e = e.parent();
                                CellAttr cellAttr = createCell();
                                cellAttr.setRowid(mainRow);
                                cellAttr.setColid(0);
                                cellAttr.setColspan(totalCols);
                                cellAttr.setFont_size("24pt");
                                cellAttr.setBold(true);
                                cellAttr.setRowspan(e.attr("rowspan").equals("") ? 1 : Integer.valueOf(e.attr("rowspan")));
                                cellAttr.setEtype(1); // 2,字段名；3,表单内容
                                cellAttr.setEvalue(e.text());
                                cellAttr.setFont_color(e.attr("color"));
                                getMainMergedInfo(mainRow, mainCol, 1, totalCols, "colspan");

                                LinkedHashMap map = new LinkedHashMap();
                                LinkedHashMap map2 = new LinkedHashMap();
                                parseSheet.buildDataEcMap(cellAttr, map);
                                parseSheet.buildPluginCellMap(cellAttr, map2);
                                pluginData.put("" + mainCol, map2);
                                mainList.add(map);
                                pluginMaintableData.put(mainRow + "", pluginData);
                                emaintable.put("ec", mainList);
                            }
                        } else if (mainList.size() > 0 && !((LinkedHashMap) mainList.get(0)).get("evalue").toString().contains(e.text())) {
                            CellAttr cellAttr = createCell();
                            if (e.select("font").size() > 0) {
                                cellAttr.setFont_size(e.selectFirst("font").attr("size"));
                            } else {
                                cellAttr.setFont_size("12pt");
                            }
                            getCoordinate();
                            cellAttr.setBold(true);
                            cellAttr.setRowid(mainRow);
                            cellAttr.setColid(mainCol);
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
                            pluginMaintableData.put(mainRow + "", pluginData);
                            emaintable.put("ec", mainList);
                        }
                        // 否则就为普通的td
                    } else if (e.select("strong").size() == 0
                            && e.getElementsByClass("InputStyle").size() == 0
                            && e.parent().select("button").size() == 0
                            && e.children().select("div[id*=button]").size() == 0) {
                        CellAttr cellAttr = createCell();
                        if (e.getElementsByClass("Label").size() > 0) {
                            cellAttr.setFieldid(e.getElementsByClass("Label").attr("name").substring(5));
                        } else {
                            Element nextElement = e.nextElementSibling();
                            if (nextElement != null) {
                                Elements inputs = nextElement.getElementsByClass("InputStyle");
                                if (inputs.size() > 0) {
                                    for (int j = 0; j < inputs.size(); j++) {
                                        cellAttr.setFieldid(inputs.get(0).attr("name").substring(5));
                                    }
                                }
                            }
                        }
                        cellAttr.setEvalue(e.text());
                        if (e.select("a").size() > 0) {
                            cellAttr.setEtype(11);
                        } else if (e.select("p").size() > 0) {
                            cellAttr.setEtype(1);
                            String text = "";
                            for (Element p : e.select("p")) {
                                text += p.text() + "\n";
                            }
                            if (e.selectFirst("p").select("font").size() > 0
                                    && e.selectFirst("p").selectFirst("font").hasAttr("color")) {
                                cellAttr.setFont_color(e.selectFirst("p").selectFirst("font").attr("color"));
                                cellAttr.setHalign(0);
                            }
                            cellAttr.setEvalue(text);
                        } else {
                            if (cellAttr.getFieldid().equals("")) {
                                cellAttr.setEtype(1);
                            } else {
                                cellAttr.setEtype(2);
                            }
                            cellAttr.setEvalue(e.text());
                        }
                        getCoordinate();
                        cellAttr.setRowid(mainRow);
                        cellAttr.setColid(mainCol);
                        cellAttr.setBackground_color("#e7f3fc");
                        getMainTableSpans(cellAttr, e);

                        LinkedHashMap map = new LinkedHashMap();
                        LinkedHashMap map2 = new LinkedHashMap();
                        parseSheet.buildDataEcMap(cellAttr, map);
                        parseSheet.buildPluginCellMap(cellAttr, map2);
                        pluginData.put("" + mainCol, map2);
                        mainList.add(map);
                        pluginMaintableData.put(mainRow + "", pluginData);
                        emaintable.put("ec", mainList);

                        String colspan = e.attr("colspan");
                        if (!colspan.equals("")) {
                            mainCol += Integer.valueOf(colspan);
                        } else {
                            mainCol++;
                        }
                        if (mainCol >= totalCols) {
                            mainCol = 0;
                        }
                    }
                }
            }
        }
        for (Element child : element.children()) {
            recursiveDFS(child);
        }
        return datajson;
    }


    /**
     * 单独处理明细表
     */
    private LinkedHashMap detailTableAnalysis(Element element) {
        LinkedHashMap detailMap = new LinkedHashMap();
        if (element.select("strong").size() > 0) {
        } else {
            int detailColCount = 0;
            int firstInputIndex = 0;
            pluginDetailtableData = new LinkedHashMap<>();
            Elements trs = element.select("tr");
            for (int i = 0; i < trs.size(); i++) {
                if (trs.get(i).select("input[class=InputStyle]").size() > 0) {
                    firstInputIndex = i;
                    break;
                }
            }
            for (int currentTrIndex = 0; currentTrIndex < trs.size(); currentTrIndex++) {
                Element e = trs.get(currentTrIndex);
                LinkedHashMap pluginData = new LinkedHashMap();
                int maxColumn = 0;
                if (e.childNodeSize() <= 1) {
                    continue;
                }
                for (Element detailtd : e.select("td")) {
                    if (checkForDisplayNone(detailtd)) {
                        continue;
                    }
                    if (detailtd.getElementsByClass("InputStyle").size() > 0) {
                        // 表单字段
                        Elements inputs = detailtd.getElementsByClass("InputStyle");
                        if (detailCol == 0) {
                            CellAttr cellAttr = createCell();
                            cellAttr.setRowid(detailRow);
                            cellAttr.setColid(detailCol);
                            cellAttr.setEtype(21);
                            cellAttr.setEvalue("选中");

                            LinkedHashMap map = new LinkedHashMap();
                            LinkedHashMap map2 = new LinkedHashMap();
                            detailList.add(map);
                            detailMap.put("ec", detailList);
                            parseSheet.buildDataEcMap(cellAttr, map);
                            parseSheet.buildPluginCellMap(cellAttr, map2);
                            detailColheads.put("col_" + detailCol, "50");
                            pluginData.put("" + detailCol, map2);
                            pluginDetailtableData.put(detailRow + "", pluginData);
                            detailCol++;
                        }
                        for (int i = 0; i < inputs.size(); i++) {
                            CellAttr cellAttr = createCell();
                            String attr = inputs.get(i).val();
                            if (attr.contains("序号")) {
                                cellAttr.setEtype(22); // 22,序号
                            } else {
                                cellAttr.setEtype(3); // 2,字段名; 3,表单字段
                            }
                            cellAttr.setRowid(detailRow);
                            cellAttr.setColid(detailCol);
                            cellAttr.setColspan(1);
                            cellAttr.setRowspan(1);
                            if (inputs.get(i).hasAttr("name") && !inputs.get(i).attr("name").equals("")) {
                                cellAttr.setFieldid(inputs.get(i).attr("name").substring(5));
                            }
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
                            detailList.add(map);
                            detailMap.put("ec", detailList);
                            detailCol++;
                        }
                    } else {
                        // 否则就为普通的 td
                        if (detailMap.size() == 0) {
                            for (Element titleElement : e.select("td")) {
                                if (titleElement.hasAttr("colspan")) {
                                    detailColCount += Integer.valueOf(titleElement.attr("colspan"));
                                } else {
                                    detailColCount++;
                                }
                            }
                            if (detailtd.text().contains("序号")) maxColumn = detailColCount;
                            else maxColumn = detailColCount + 1;
                            for (int i = 0; i < detailColCount + 1; i++) {
                                CellAttr cellAttr = new CellAttr();
                                // 明细空一行的最后一格显示添加和删除按钮
                                if (i == detailColCount) {
                                    cellAttr.setEtype(10);
                                    cellAttr.setBright_style(1);
                                    cellAttr.setBright_color("#90badd");
                                    cellAttr.setBackground_image("/workflow/exceldesign/image/shortBtn/detail/de_btn_wev8.png");
                                    cellAttr.setBackground_imagelayout("3");
                                }
                                cellAttr.setRowid(detailRow);
                                cellAttr.setColid(detailCol);
                                LinkedHashMap map = new LinkedHashMap();
                                LinkedHashMap map2 = new LinkedHashMap();
                                detailList.add(map);
                                detailMap.put("ec", detailList);
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
                        if (detailCol == 0 && currentTrIndex < firstInputIndex) {
                            CellAttr cellAttr = createCell();
                            cellAttr.setRowid(detailRow);
                            cellAttr.setColid(detailCol);
                            cellAttr.setEvalue("全选");
                            cellAttr.setEtype(20);
                            cellAttr.setColspan(1);
                            cellAttr.setBackground_image("/workflow/exceldesign/image/shortBtn/detail/de_checkall_wev8.png");
                            cellAttr.setBackground_imagelayout("3");
                            if (firstInputIndex - currentTrIndex > 1) {
                                cellAttr.setRowspan(firstInputIndex - currentTrIndex);
                                getDetailMergedInfo(detailRow, detailCol, cellAttr.getRowspan(), "rowspan");
                            }
                            LinkedHashMap map = new LinkedHashMap();
                            LinkedHashMap map2 = new LinkedHashMap();
                            detailList.add(map);
                            detailMap.put("ec", detailList);
                            detailColheads.put("col_" + detailCol, "50");
                            parseSheet.buildDataEcMap(cellAttr, map);
                            parseSheet.buildPluginCellMap(cellAttr, map2);
                            pluginData.put("" + detailCol, map2);
                            detailCol++;
                        }
                        CellAttr cellAttr = createCell();
                        cellAttr.setRowid(detailRow);
                        cellAttr.setColid(detailCol);
                        cellAttr.setEvalue(detailtd.text());

                        getDetailTableSpans(cellAttr, detailtd);

                        if (detailtd.getElementsByClass("Label").size() > 0) {
                            cellAttr.setFieldid(detailtd.getElementsByClass("Label").attr("name").substring(5));
                        } else {
                            Element firstInputTr = trs.get(firstInputIndex);
                            if (currentTrIndex == firstInputIndex) {
                                if (detailtd.hasAttr("name") && detailtd.attr("name").contains("field")) {
                                    cellAttr.setFieldid(detailtd.attr("name").substring(5));
                                }
                            } else if (currentTrIndex == firstInputIndex - 1) {
                                if (firstInputTr.select("td").size() > detailCol - 1) {
                                    Elements es = firstInputTr.select("td").get(detailCol - 1).getElementsByClass("InputStyle");
                                    if (es.size() > 0) {
                                        cellAttr.setFieldid(es.attr("name").substring(5));
                                    }
                                }
                            } else if (currentTrIndex > firstInputIndex) {
                                if (detailtd.nextElementSibling() != null) {
                                    Element specialCell = detailtd.nextElementSibling().selectFirst("input[class=InputStyle]");
                                    if (specialCell != null && specialCell.hasAttr("name")) {
                                        cellAttr.setFieldid(specialCell.attr("name").substring(5));
                                    }
                                }
                            }
                        }
                        if (cellAttr.getFieldid().equals("")) {
                            cellAttr.setEtype(1);
                        } else {
                            cellAttr.setEtype(2);
                        }
                        if (cellAttr.getEvalue().contains("序号")) detailColheads.put("col_" + detailCol, "50");
                        else detailColheads.put("col_" + detailCol, "120");
                        LinkedHashMap map = new LinkedHashMap();
                        LinkedHashMap map2 = new LinkedHashMap();
                        parseSheet.buildDataEcMap(cellAttr, map);
                        parseSheet.buildPluginCellMap(cellAttr, map2);
                        pluginData.put("" + detailCol, map2);
                        detailList.add(map);
                        detailMap.put("ec", detailList);
                        detailCol++;
                    }
                    int index = (detailtd.siblingIndex() + 1) / 2 - 1;
                    pluginDetailtableData.put(detailRow + "", pluginData);
                    if (detailtd.nextElementSibling() == null
                            || (checkForDisplayNone(detailtd.nextElementSibling())
                            && e.select("td:gt(" + index + ")").size() == e.select("td[style=display: none]:gt(" + index + ")").size())) {
                        maxColumn = detailCol > maxColumn ? detailCol : maxColumn;
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
                if (e.nextElementSibling() == null) {
                    cellAttr.setEtype(9);
                    cellAttr.setEvalue("表尾标识");
                    cellAttr.setValign(1);
                    cellAttr.setBackground_color("#eeeeee");
                    detailMap.put("edtailinrow", detailRow + "");
                    LinkedHashMap mergedCell = new LinkedHashMap();
                    mergedCell.put("row", detailRow);
                    mergedCell.put("rowCount", 1);
                    mergedCell.put("col", detailCol);
                    mergedCell.put("colCount", detailColCount + 1);
                    detailSpans.add(mergedCell);
                } else if (currentTrIndex == firstInputIndex - 1) {
                    cellAttr.setEtype(8);
                    cellAttr.setEvalue("表头标识");
                    cellAttr.setValign(1);
                    cellAttr.setBackground_color("#eeeeee");
                    detailMap.put("edtitleinrow", detailRow + "");
                    LinkedHashMap mergedCell = new LinkedHashMap();
                    mergedCell.put("row", detailRow);
                    mergedCell.put("rowCount", 1);
                    mergedCell.put("col", detailCol);
                    mergedCell.put("colCount", detailColCount + 1);
                    detailSpans.add(mergedCell);
                }
                if (cellAttr.getEtype() != 0) {
                    pluginData.put(detailCol, map2);
                    cellAttr.setRowid(detailRow);
                    cellAttr.setColid(detailCol);
                    parseSheet.buildDataEcMap(cellAttr, map);
                    parseSheet.buildPluginCellMap(cellAttr, map2);
                    detailList.add(map);
                    detailMap.put("ec", detailList);
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
     * 构造明细表的 pluginjson
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

        LinkedHashMap<String, Object> dataTable = new LinkedHashMap<String, Object>();
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
            rowHeaderData.put("rowCount", mainRow >= 20 ? mainRow + 5 : 20);
            colHeaderData.put("colCount", totalCols >= 15 ? totalCols + 3 : 15);
            rowHeaderData.put("defaultDataNode", defaultDataNode);
            colHeaderData.put("defaultDataNode", defaultDataNode);
            rowRangeGroup.put("itemsCount", mainRow >= 20 ? mainRow + 5 : 20);
            colRangeGroup.put("itemsCount", totalCols >= 15 ? totalCols + 3 : 15);
            Sheet1.put("rowRangeGroup", rowRangeGroup);
            Sheet1.put("colRangeGroup", colRangeGroup);
            Sheet1.put("spans", mainSpans);
            dataTable.put("rowCount", mainRow >= 20 ? mainRow + 5 : 20);
            dataTable.put("colCount", totalCols >= 15 ? totalCols + 3 : 15);
        } else if (str.contains("detail")) {
            int colCount = ((LinkedHashMap) data.get("0")).size();
            int rowCount = data.size();
            if (data.size() > 3) {
                colCount = ((LinkedHashMap) data.get("3")).size();
            }
            rowHeaderData.put("rowCount", rowCount >= 10 ? rowCount + 5 : 10);
            colHeaderData.put("colCount", colCount >= 20 ? colCount + 5 : 20);
            rowHeaderData.put("defaultDataNode", defaultDataNode);
            colHeaderData.put("defaultDataNode", defaultDataNode);
            rowRangeGroup.put("itemsCount", rowCount >= 10 ? rowCount + 5 : 10);
            colRangeGroup.put("itemsCount", colCount >= 20 ? colCount + 5 : 20);
            Sheet1.put("rowRangeGroup", rowRangeGroup);
            Sheet1.put("colRangeGroup", colRangeGroup);
            Sheet1.put("spans", detailSpans);
            dataTable.put("rowCount", rowCount >= 10 ? rowCount + 5 : 10);
            dataTable.put("colCount", colCount >= 20 ? colCount + 5 : 20);
        }
        dataTable.put("dataTable", data);
        Sheet1.put("columns", getPluginColumns(str, data));
        Sheet1.put("data", dataTable);
        pluginjson.put(str, pluginMap);
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
            int serialNumIndex = 0;
            for (int i = 0; i < map.size(); i++) {
                if (((LinkedHashMap) map.get("" + i)).size() > ((LinkedHashMap) map.get("" + (i + 1))).size()) {
                    serialNumIndex = i;
                    break;
                }
            }
            boolean isSerialNum = ((LinkedHashMap) ((LinkedHashMap) map.get("" + serialNumIndex)).get("1")).containsValue("序号");
            for (int i = 0; i < ((LinkedHashMap) map.get("" + serialNumIndex)).size(); i++) {
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
    private void getDetailTableCell(String str) {
        CellAttr cellAttr = createCell();
        cellAttr.setRowid(mainRow);
        cellAttr.setColid(mainCol);
        cellAttr.setColspan(totalCols - mainCol);
        cellAttr.setEtype(7); // 明细
        cellAttr.setEvalue(str.substring(7));

        LinkedHashMap map = new LinkedHashMap();
        parseSheet.buildDataEcMap(cellAttr, map);
        mainList.add(map);
        emaintable.put("ec", mainList);
        mainRowheads.put("row_" + mainRow, "30");
    }


    /**
     * 获取主表colspan和rowspan信息
     */
    private void getMainTableSpans(CellAttr cellAttr, Element e) {
        if (e.childNodeSize() == 1 && e.select("strong").size() > 0) {
            cellAttr.setColspan(totalCols);
            cellAttr.setRowspan(1);
            getMainMergedInfo(mainRow, 0, 1, totalCols, "colspan");
            return;
        }
        boolean hasColspan = !e.attr("colspan").equals("");
        boolean hasRowspan = !e.attr("rowspan").equals("");
        if (e.tagName().equals("td") && e.select("table[name*=oTable]").size() > 0) {
            getMainMergedInfo(mainRow, mainCol, 1, totalCols - mainCol, "colspan");
            cellAttr.setColspan(totalCols);
        } else {
            // 只有列合并 没有行合并
            if (hasColspan && !hasRowspan) {
                if (mainCol == totalCols) {
                    mainCol = 0;
                }
                if (mainCol + Integer.valueOf(e.attr("colspan")) > totalCols) {
                    getMainMergedInfo(mainRow, mainCol, 1, totalCols - mainCol, "colspan");
                    cellAttr.setColspan(totalCols - mainCol);
                } else {
                    getMainMergedInfo(mainRow, mainCol, 1, Integer.valueOf(e.attr("colspan")), "colspan");
                    cellAttr.setColspan(Integer.valueOf(e.attr("colspan")));
                }
            } else {
                cellAttr.setColspan(1);
                mainSpanArray[mainRow][mainCol] = 1;
            }
        }
        // 只有行合并 没有列合并
        if (hasRowspan && !hasColspan) {
            if (mainCol == totalCols) {
                mainCol = 0;
            }
            getMainMergedInfo(mainRow, mainCol, Integer.valueOf(e.attr("rowspan")), 1, "rowspan");
            cellAttr.setRowspan(Integer.valueOf(e.attr("rowspan")));
        } else {
            cellAttr.setRowspan(1);
            mainSpanArray[mainRow][mainCol] = 1;
        }
        // 行列合并都有
        if (hasColspan && hasRowspan) {
            getMainMergedInfo(mainRow, mainCol, Integer.valueOf(e.attr("rowspan")), Integer.valueOf(e.attr("colspan")), "colandrow");
            cellAttr.setRowspan(Integer.valueOf(e.attr("rowspan")));
            cellAttr.setColspan(Integer.valueOf(e.attr("colspan")));
        }
    }


    /**
     * 获取明细表colspan和rowspan信息
     */
    private void getDetailTableSpans(CellAttr cellAttr, Element e) {
        boolean hasColspan = !e.attr("colspan").equals("");
        boolean hasRowspan = !e.attr("rowspan").equals("");
        if (hasColspan) {
            int colspan = Integer.valueOf(e.attr("colspan"));
            getDetailMergedInfo(detailRow, detailCol, colspan, "colspan");
            cellAttr.setColspan(colspan);
            detailCol += colspan - 1;
        } else {
            cellAttr.setColspan(1);
        }
        if (hasRowspan) {
            int rowspan = Integer.valueOf(e.attr("rowspan"));
            getDetailMergedInfo(detailRow, detailCol, rowspan, "rowspan");
            cellAttr.setRowspan(rowspan);
        } else {
            cellAttr.setRowspan(1);
        }
    }


    /**
     * 获取主表合并单元格信息
     */
    private void getMainMergedInfo(int row, int col, int rowspan, int colspan, String flag) {
        if (flag.equals("colspan")) {
            LinkedHashMap mergedCell = new LinkedHashMap();
            mergedCell.put("row", row);
            mergedCell.put("rowCount", 1);
            mergedCell.put("col", col);
            mergedCell.put("colCount", colspan);
            mainSpans.add(mergedCell);
            for (int i = 0; i < colspan; i++) {
                mainSpanArray[row][col + i] = 1;
            }
        } else if (flag.equals("rowspan")) {
            LinkedHashMap mergedCell = new LinkedHashMap();
            mergedCell.put("row", row);
            mergedCell.put("rowCount", rowspan);
            mergedCell.put("col", col);
            mergedCell.put("colCount", 1);
            mainSpans.add(mergedCell);
            for (int i = 0; i < rowspan; i++) {
                mainSpanArray[row + i][col] = 1;
            }
        } else if (flag.equals("colandrow")) {
            LinkedHashMap mergedCell = new LinkedHashMap();
            mergedCell.put("row", row);
            mergedCell.put("rowCount", rowspan);
            mergedCell.put("col", col);
            mergedCell.put("colCount", colspan);
            mainSpans.add(mergedCell);
            for (int i = 0; i < rowspan; i++) {
                for (int j = 0; j < colspan; j++) {
                    mainSpanArray[mainRow + i][mainCol + j] = 1;
                }
            }
        }
    }


    /**
     * 获取明细合并单元格信息
     */
    private void getDetailMergedInfo(int row, int col, int span, String flag) {
        if (flag.equals("colspan")) {
            LinkedHashMap mergedCell = new LinkedHashMap();
            mergedCell.put("row", row);
            mergedCell.put("rowCount", 1);
            mergedCell.put("col", col);
            mergedCell.put("colCount", span);
            detailSpans.add(mergedCell);
        } else if (flag.equals("rowspan")) {
            LinkedHashMap mergedCell = new LinkedHashMap();
            mergedCell.put("row", row);
            mergedCell.put("rowCount", span);
            mergedCell.put("col", col);
            mergedCell.put("colCount", 1);
            detailSpans.add(mergedCell);
        }
    }


    /**
     * 检查元素是否含有属性 style=display:none
     */
    private boolean checkForDisplayNone(Element element) {
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
        cellAttr.setHalign(1); // 水平居中
        cellAttr.setValign(1); // 竖直居中
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


    /**
     * 根据模板坐标数据确定此单元格的位置，并进行异常处理
     * mainSpanArray 的结构为如下，当对应的行列上有元素时，行列作为脚标对应的值为 1
     * row\col    0   1   2   3   4   ...
     *  0         1   1   1   1   1
     *  1         1   1   1   0   0
     *  2         1   1   1   1   1
     *  3         1   1   1   1   1
     *  4         1   1   1   1   1
     *  ...
     */
    private void getCoordinate() {
        while (mainSpanArray[mainRow][mainCol] == 1) {
            mainCol++;
            if (mainCol == totalCols) {
                mainRow++;
                mainCol = 0;
            }
        }
    }
}

