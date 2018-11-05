import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

public class Attr {

    // 获取整个布局的宽度
    public Integer getMainTableWidth(Document document) {
        // 不会出现百分比的情况，不特殊处理
        return Integer.valueOf(document.getElementsByTag("table").first().attr("width"));
    }

    // 获取整个布局的border
    public Integer getMainTableBorder(Document document) {
        String border = document.getElementsByTag("table").first().attr("border").trim();
        if ("".equals(border)) {
            return 0;
        }
        return Integer.valueOf(border);
    }

    // 获取整个布局的padding
    public Integer getMainTablePadding(Document document) {
        String padding = document.getElementsByTag("table").first().attr("cellpadding").trim();
        if ("".equals(padding)) {
            return 0;
        }
        return Integer.valueOf(padding);
    }

//    // 获取明细表的宽度
//    public Integer getDetailTableWidth(Elements elements, Integer n) {
//        Boolean b  = elements.get(n).attr("style").split(":")[1].endsWith("%");
//        if (b) {
//            return Integer.valueOf(elements.get(n).attr("style").split(":")[1].split("%")[0].trim());
//        }
////        return Integer.valueOf(elements.get(n).attr("style").split(":")[1].trim());
//
//        return b ?
//                Integer.valueOf(elements.get(n).attr("style").split(":")[1].split("%")[0].trim()) :
//                Integer.valueOf(elements.get(n).attr("style").split(":")[1].trim());
//    }
}
