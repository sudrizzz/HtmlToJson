import java.util.*;

/**
 * 新表单设计器Excel模板导入-单个Sheet解析
 * @author liuzy 2015-02-02
 */
public class ParseSheet{
    private int detailflag;
    private Map<String,Object> fieldmap = new HashMap<String,Object>();

    private int rownum;
    private int colnum;
    private List<String> detailSheet = new ArrayList<String>();
    private int edtitleinrow = -1;
    private int edtailinrow = -1;
    private Map<String,String> fieldattr_map = new HashMap<String,String>();

    private Map<String,Object> data_rowheads = new LinkedHashMap<String,Object>();
    private Map<String,Object> data_colheads = new LinkedHashMap<String,Object>();
    private List<LinkedHashMap<String,Object>> data_ec = new ArrayList<LinkedHashMap<String,Object>>();

    private List<LinkedHashMap<String,Object>> plugin_rows = new ArrayList<LinkedHashMap<String,Object>>();
    private List<LinkedHashMap<String,Object>> plugin_columns = new ArrayList<LinkedHashMap<String,Object>>();
    private List<LinkedHashMap<String,Integer>> plugin_combine = new ArrayList<LinkedHashMap<String,Integer>>();
    private Map<String,Object> plugin_data = new LinkedHashMap<String,Object>();



    public void buildDataEcMap(CellAttr cellattr,LinkedHashMap<String,Object> ec_map){
        ec_map.put("id", cellattr.getRowid()+","+cellattr.getColid());
        ec_map.put("rowspan", cellattr.getRowspan()+"");
        ec_map.put("colspan", cellattr.getColspan()+"");
        ec_map.put("etype", cellattr.getEtype()+"");
        ec_map.put("evalue", cellattr.getEvalue());

        if(cellattr.getEtype()==8||cellattr.getEtype()==9)	return;
        if(cellattr.getEtype()==7)
            ec_map.put("detail", "detail_"+cellattr.getEvalue().replace("明细表", "").replace("明细", ""));

        ec_map.put("field", cellattr.getFieldid()+"");
        ec_map.put("fieldtype", cellattr.getFieldtype());
        if(!"".equals(cellattr.getBackground_color()))
            ec_map.put("backgroundColor", cellattr.getBackground_color());
        LinkedHashMap<String,Object> font_map = new LinkedHashMap<String,Object>();
        ec_map.put("font", font_map);
        if(cellattr.isItalic())		font_map.put("italic", "true");
        if(cellattr.isBold())		font_map.put("bold", "true");
        if(cellattr.isDeleteline())	font_map.put("deleteline", "true");
        if(cellattr.isUnderline())	font_map.put("underline", "true");
        switch(cellattr.getHalign()){
            case 0:	font_map.put("text-align", "left");	break;
            case 1: font_map.put("text-align", "center"); break;
            case 2: font_map.put("text-align", "right"); break;
        }
        switch(cellattr.getValign()){
            case 0: font_map.put("valign", "top"); break;
            case 1: font_map.put("valign", "middle"); break;
            case 2: font_map.put("valign", "bottom"); break;
        }
        if(cellattr.isWordwrap())	font_map.put("autoWrap", "true");
        font_map.put("font-size", cellattr.getFont_size());
        font_map.put("font-family", cellattr.getFont_family());
        font_map.put("color", cellattr.getFont_color());

        if(cellattr.getIndent()>0)
            ec_map.put("etxtindent", String.valueOf(cellattr.getIndent()));

        List<LinkedHashMap<String,Object>> borderList = new ArrayList<LinkedHashMap<String,Object>>();
        ec_map.put("eborder", borderList);
        LinkedHashMap<String,Object> top_border = new LinkedHashMap<String,Object>();
        top_border.put("kind", "top");
        top_border.put("style", cellattr.getBtop_style());
        top_border.put("color", cellattr.getBtop_color());
        borderList.add(top_border);
        LinkedHashMap<String,Object> bottom_border = new LinkedHashMap<String,Object>();
        bottom_border.put("kind", "bottom");
        bottom_border.put("style", cellattr.getBbottom_style());
        bottom_border.put("color", cellattr.getBbottom_color());
        borderList.add(bottom_border);
        LinkedHashMap<String,Object> left_border = new LinkedHashMap<String,Object>();
        left_border.put("kind", "left");
        left_border.put("style", cellattr.getBleft_style());
        left_border.put("color", cellattr.getBleft_color());
        borderList.add(left_border);
        LinkedHashMap<String,Object> right_border = new LinkedHashMap<String,Object>();
        right_border.put("kind", "right");
        right_border.put("style", cellattr.getBright_style());
        right_border.put("color", cellattr.getBright_color());
        borderList.add(right_border);
    }

    public void buildPluginCellMap(CellAttr cellattr,LinkedHashMap<String,Object> cell_map){
        cell_map.put("value", cellattr.getEvalue());
        LinkedHashMap<String,Object> cell_style_map = new LinkedHashMap<String,Object>();
        cell_map.put("style", cell_style_map);

        if(!"".equals(cellattr.getBackground_color()))
            cell_style_map.put("backColor", cellattr.getBackground_color());
        if(cellattr.getEtype()==3){		//字段
            cell_style_map.put("textIndent", cellattr.getIndent()+2.5);
            String field_png_path = "/workflow/exceldesign/image/controls/"+cellattr.getFieldtype()+cellattr.getFieldattr()+"_wev8.png";
            cell_style_map.put("backgroundImage", field_png_path);
            cell_style_map.put("backgroundImageLayout", 3);
        }else if(cellattr.getEtype()==7){		//明细表
            cell_style_map.put("backgroundImage", "/workflow/exceldesign/image/shortBtn/detail/detailTable_wev8.png");
            cell_style_map.put("backgroundImageLayout", 3);
            cell_style_map.put("textIndent", 3);
        }else if(cellattr.getEtype()==10){		//明细按钮
            cell_style_map.put("backgroundImage", "/workflow/exceldesign/image/shortBtn/detail/de_btn_wev8.png");
            cell_style_map.put("backgroundImageLayout", 3);
        }else if(cellattr.getEtype()==8||cellattr.getEtype()==9){		//表头、表尾
            return;
        }else{
            if(cellattr.getIndent()>0)
                cell_style_map.put("textIndent", cellattr.getIndent());
        }

        String font_style = "";
        if(cellattr.isItalic())	font_style += "italic ";
        if(cellattr.isBold())		font_style += "bold ";
        font_style += cellattr.getFont_size()+" ";
        font_style += cellattr.getFont_family();
        cell_style_map.put("font", font_style);
        if(!"".equals(cellattr.getFont_color()))
            cell_style_map.put("foreColor", cellattr.getFont_color());

        cell_style_map.put("wordWrap", cellattr.isWordwrap());
        if(cellattr.isUnderline()&&cellattr.isDeleteline()){
            cell_style_map.put("textDecoration", 3);
        }else if(cellattr.isUnderline()){
            cell_style_map.put("textDecoration", 1);
        }else if(cellattr.isDeleteline()){
            cell_style_map.put("textDecoration", 2);
        }
        cell_style_map.put("hAlign", cellattr.getHalign());
        cell_style_map.put("vAlign", cellattr.getValign());

        LinkedHashMap<String,Object> borderLeft = new LinkedHashMap<String,Object>();
        if(!"".equals(cellattr.getBleft_color()))
            borderLeft.put("color", cellattr.getBleft_color());
        borderLeft.put("style", cellattr.getBleft_style());
        cell_style_map.put("borderLeft", borderLeft);

        LinkedHashMap<String,Object> borderRight = new LinkedHashMap<String,Object>();
        if(!"".equals(cellattr.getBright_color()))
            borderRight.put("color", cellattr.getBright_color());
        borderRight.put("style", cellattr.getBright_style());
        cell_style_map.put("borderRight", borderRight);

        LinkedHashMap<String,Object> borderTop = new LinkedHashMap<String,Object>();
        if(!"".equals(cellattr.getBtop_color()))
            borderTop.put("color", cellattr.getBtop_color());
        borderTop.put("style", cellattr.getBtop_style());
        cell_style_map.put("borderTop", borderTop);

        LinkedHashMap<String,Object> borderBottom = new LinkedHashMap<String,Object>();
        if(!"".equals(cellattr.getBbottom_color()))
            borderBottom.put("color", cellattr.getBbottom_color());
        borderBottom.put("style", cellattr.getBbottom_style());
        cell_style_map.put("borderBottom", borderBottom);
    }


    public int getDetailflag() {
        return detailflag;
    }

    public Map<String, Object> getFieldmap() {
        return fieldmap;
    }

    public int getRownum() {
        return rownum;
    }

    public int getColnum() {
        return colnum;
    }

    public List<String> getDetailSheet() {
        return detailSheet;
    }

    public int getEdtitleinrow() {
        return edtitleinrow;
    }

    public int getEdtailinrow() {
        return edtailinrow;
    }

    public Map<String, String> getFieldattr_map() {
        return fieldattr_map;
    }

    public Map<String, Object> getData_rowheads() {
        return data_rowheads;
    }

    public Map<String, Object> getData_colheads() {
        return data_colheads;
    }

    public List<LinkedHashMap<String, Object>> getData_ec() {
        return data_ec;
    }

    public List<LinkedHashMap<String, Object>> getPlugin_rows() {
        return plugin_rows;
    }

    public List<LinkedHashMap<String, Object>> getPlugin_columns() {
        return plugin_columns;
    }

    public List<LinkedHashMap<String, Integer>> getPlugin_combine() {
        return plugin_combine;
    }

    public Map<String, Object> getPlugin_data() {
        return plugin_data;
    }


}