import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;

import org.json.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class Main {

    public static void main(String[] args) {
        String filepath = "layout.html";
//        Table table = new Table();
//        table.get(filepath, "GBK");

        Cell cell = new Cell();
        cell.getTable(filepath, "GBK");

//        GetJson getJson = new GetJson();
//        getJson.get(filepath, "GBK"); // 防止读出来乱码

    }

//    public static String get(String filepath, String encoding) {
//
//        Attr attr = new Attr();
//
//        try {
//            // 读取文件将内容转成字符串
//            File file = new File(filepath);
//            FileInputStream fin = new FileInputStream(filepath);
//            InputStreamReader reader = new InputStreamReader(fin, encoding);
//            BufferedReader buffReader = new BufferedReader(reader);
//            StringBuffer sb = new StringBuffer();
//            String strTmp = "";
//            while ((strTmp = buffReader.readLine()) != null) {
//                sb.append(strTmp);
//            }
//
//            // 使用Jsoup解析html字符串
//            Document document = Jsoup.parse(sb.toString());
//
//            // 获取整个表单的宽度，border，padding
//            Integer width = attr.getMainTableWidth(document);
//            Integer border = attr.getMainTableBorder(document);
//            Integer padding = attr.getMainTablePadding(document);
//
//            Elements elements1 = document.getElementsByClass("ListStyle detailtable detailtableTopTable"); // 明细表
//            Elements elements2 = document.select("table").select("font"); // 单独设置字体属性的是表名，先主表后明细表
//            Elements elements3 = document.select("table").select("table"); // 明细表元素
//            Elements elements4 = document.select("table").select("tr").select("td"); // 主表每一行每一列（每一个字段）
//
//            HashMap layout = new HashMap();
//            HashMap eformdesign = new HashMap(); // 整个布局
//            HashMap eattr = new HashMap(); // 表单基本信息
//            HashMap etables = new HashMap(); // 主表和明细表内容
//            HashMap detailtable = new HashMap(); // 明细表信息
//            HashMap emaintable = new HashMap(); // 明细表信息
//            HashMap ecmain = new HashMap(); // 主表每个字段的信息
//            ArrayList ecdt = new ArrayList(); // 明细表每个字段的信息
//
//            // 循环获取明细表信息
//            // key相同，会把之前的值覆盖掉，但是每个明细表都会有rowheads和colheads...
//            // 所以暂时先处理最后一个明细表，这个布局里有两个，暂时处理第一个
//            for (int m = 0; m < elements1.first().select("tr").size(); m++) {
//                HashMap rowheadsdt = new HashMap(); // 行信息
//                HashMap colheadsdt = new HashMap(); // 列信息
//
//                // 获取明细表的宽度，border，padding
//                Integer widthdt = 0;
//                if (elements1.get(m).attr("style").split(":")[1].endsWith("%")) { // 如果是百分比形式，进行特殊处理
//                    widthdt = Integer.valueOf(elements1.get(m).attr("style").split(":")[1].split("%")[0].trim()) / 100 * width - border * 2 - padding * 2;
//                } else {
//                    widthdt = Integer.valueOf(elements1.get(m).attr("style").split(":")[1].trim()); // 不是百分比形式，直接取值
//                }
//                Integer borderdt = Integer.valueOf(elements1.get(m).attr("border").trim().equals("") ? "0" : elements1.get(m).attr("border").trim());
//                Integer paddingdt = Integer.valueOf(elements1.get(m).attr("padding").trim().equals("") ? "0" : elements1.get(m).attr("padding").trim());
//                Elements midTable1td1 = elements1.get(m).select("td.midTable1td1");
//                Elements midTable1td2 = elements1.get(m).select("td.midTable1td2");
//                Integer colwidthdt = (widthdt - borderdt * 2 - paddingdt * 2) / midTable1td1.size();
//
//                // 每一列
//                String colspan = ""; // 行单元格合并
//                String rowspan = ""; // 列单元格合并
//                for (int n = 0; n < midTable1td1.size(); n++) {
//
//                    HashMap cell = new HashMap();
//                    rowheadsdt.put("row_" + n, 30); // 顺序不是按照put的先后顺序排列...
//                    colheadsdt.put("col_" + n, colwidthdt);
//
//                    // 对colspan和rowspan之后的单元格定位进行处理
//                    cell.put("id", m + "," + (n + (colspan.equals("") ? 0 : Integer.valueOf(colspan))));
//                    colspan = elements1.select("tr").select("td").get(m * n + n).attr("colspan").trim();
//
//                    // 列合并比较复杂，先不考虑，假设所有单元格都没有列合并
////                    rowspan = elements1.select("tr").select("td").get(m * n + n).attr("rowspan").trim();
//                    cell.put("colspan", colspan.equals("") ? 0 : Integer.valueOf(colspan));
//                    cell.put("rowspan", 0);
//                    cell.put("etype", elements1.first().select("td").attr("td.InputStyle"));
//                    ecdt.add(n, cell);
//                }
////                for (int n = 0; n < midTable1td2.size(); n++) {
////                    rowheadsdt.put("row_" + n, 30);
////                    colheadsdt.put("col_" + n, colwidthdt);
////                }
//
//                detailtable.put("rowheads", rowheadsdt);
//                detailtable.put("colheads", colheadsdt);
//                detailtable.put("ec", ecdt);
//                etables.put("detail_" + (m + 1), detailtable);
//
//            }
//
//            // 将主表和明细表的信息放到总表单内
//            etables.put("emaintable", emaintable);
//            eattr.put("formname", file.getName());
//            eformdesign.put("eattr", eattr);
//            eformdesign.put("etables", etables);
//            layout.put("eformdesign", eformdesign);
//
//            System.out.println(new JSONObject(layout));
//
//            buffReader.close();
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//        return null;
//    }
}
