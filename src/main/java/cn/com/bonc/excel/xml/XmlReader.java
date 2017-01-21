/*****************************************************************************
 * 东方国信Excel助手[ExcelHelper]
 *----------------------------------------------------------------------------
 * cn.com.bonc.excel.xml.XmlReader.java
 *
 * @author andy
 * @date 2017年1月17日
 * @version 0.0.1
 * @since 0.0.1
 *----------------------------------------------------------------------------
 * (C) 北京东方国信科技股份有限公司
 *     Business-intelligence Of Oriental Nations Corporation Ltd. 2016
 *****************************************************************************/
package cn.com.bonc.excel.xml;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import cn.com.bonc.excel.util.EmptyUtil;
import cn.com.bonc.excel.xml.bean.CellBean;
import cn.com.bonc.excel.xml.bean.DataBean;
import cn.com.bonc.excel.xml.bean.RowBean;
import cn.com.bonc.excel.xml.bean.SheetBean;
import cn.com.bonc.excel.xml.bean.StyleBean;

/**
 * Xml解析类
 * cn.com.bonc.excel.xml.XmlReader.java
 * 
 * @author andy
 * @date 2017年1月17日
 *
 * @since 0.0.1
 */
public class XmlReader {
	/**
	 * 解析xml获取sheet页数据集list
	 * 
	 * @param fileName xml文件名
	 * @param params 入参map
	 * @return List<SheetBean> sheet页数据集list
	 * @author andy
	 * @throws Exception
	 * @date 2017年1月18日
	 */
    public static List<SheetBean> readXml(String fileName, Map<String, Object> params) throws Exception {
        InputStream inputStream = XmlReader.class.getClassLoader().getResourceAsStream(fileName);
        return readXml(inputStream, params);
    }

    /**
     * 解析xml获取sheet页数据集list
     * 
     * @param inputStream xml文件流
     * @param params 入参map
     * @return List<SheetBean> sheet页数据集list
     * @author andy
     * @throws Exception
     * @date 2017年1月18日
     */
    @SuppressWarnings("unchecked")
    public static List<SheetBean> readXml(InputStream inputStream, Map<String, Object> params) throws Exception {
        // 返回结果sheet页数据集list
        List<SheetBean> sheetList = new ArrayList<SheetBean>();
        // 定义解析器
        SAXReader saxReader = new SAXReader();
        // 解析xml为dom对象
        Document document = saxReader.read(inputStream);
        // 获取根节点root
        Element root = document.getRootElement();
        // 获取全部sheet节点
        List<Element> sheets = root.elements("SHEET");
        // 定义sheet页数据集
        SheetBean sheetBean = null;
        // 遍历sheet节点集合
        for (Element sheet : sheets) {
            // 获取sheet页数据集
            sheetBean = getSheet(sheet, params);
            // 添加sheet页数据集
            sheetList.add(sheetBean);
        }
        // 返回结果
        return sheetList;
    }

    /**
     * 获取sheet页数据集
     * 
     * @param sheet sheet节点
     * @param params 入参map
     * @return SheetBean
     * @author andy
     * @throws Exception
     * @date 2017年1月18日
     */
    private static SheetBean getSheet(Element sheet, Map<String, Object> params) throws Exception {
        // 返回结果sheet页数据集
    	SheetBean sheetBean = new SheetBean();
        // 获取sheet页名称
        String sheetName = sheet.attributeValue("name");
        // 设置sheet页名称
        if(EmptyUtil.isNotEmpty(sheetName)){
        	sheetBean.setSheetName(sheetName);
        }
        // 获取单元格样式map
        Map<String, StyleBean> styles = getStyles(sheet);
        // 设置单元格样式map
        sheetBean.setStyles(styles);
        // 获取标题节点
        Element title = sheet.element("TITLE");
    	// 获取表头数据
        List<RowBean> rows = getRows(title, params);
    	// 设置表头数据
        sheetBean.setTitle(rows);
        // 获取填充数据区
        DataBean dataBean = getData(sheet, params);
        // 设置填充数据区
        sheetBean.setData(dataBean);
        // 返回结果
        return sheetBean;
    }

    /**
     * 获取单元格样式map
     * 
     * @param sheet sheet节点
     * @return Map<String, StyleBean> 
     * @author andy
     * @throws Exception
     * @date 2017年1月18日
     */
    @SuppressWarnings("unchecked")
	private static Map<String, StyleBean> getStyles(Element sheet) throws Exception {
        // 返回结果单元格样式map
        Map<String, StyleBean> styleMap = new HashMap<String, StyleBean>();
        // 获取全部单元格样式节点
        List<Element> styles = sheet.elements("STYLE");
        // 单元格样式
        StyleBean styleBean = null;
        // 遍历单元格样式节点集合
        for (Element style : styles) {
            // 创建单元格样式
            styleBean = new StyleBean();
            // 样式id
            String styleId = style.attributeValue("id");
            if(EmptyUtil.isNotEmpty(styleId)){
            	// 背景颜色
                String bgColor = style.attributeValue("bgcolor");
                // 判断背景颜色
                if(EmptyUtil.isNotEmpty(bgColor)){
                	// 设置背景颜色
                    styleBean.setBackgroundColor(Short.parseShort(bgColor));
                }
                // 字体大小
                String fontSize = style.attributeValue("fontsize");
                // 判断字体大小
                if(EmptyUtil.isNotEmpty(fontSize)){
                	// 设置字体大小
                    styleBean.setFontSize(Short.parseShort(fontSize));
                }
                // 字体颜色
                String fontColor = style.attributeValue("fontcolor");
                // 判断字体颜色
                if(EmptyUtil.isNotEmpty(fontColor)){
                	// 设置字体颜色
                    styleBean.setFontColor(Short.parseShort(fontColor));
                }
                // 加粗值
                String boldWeight = style.attributeValue("boldweight");
                // 判断加粗值
                if(EmptyUtil.isNotEmpty(boldWeight)){
                	// 设置加粗值
                    styleBean.setBoldWeight(Short.parseShort(boldWeight));
                }
                // 对齐方式
                String align = style.attributeValue("align");
                // 判断对齐方式
                if(EmptyUtil.isNotEmpty(align)){
                	// 设置对齐方式
                    styleBean.setAlign(Short.parseShort(align));
                }
                // 是否有边框
                String border = style.attributeValue("border");
                // 判断是否有边框
                if(EmptyUtil.isNotEmpty(border)){
                	// 设置是否有边框
                    styleBean.setBorder(Boolean.parseBoolean(border));
                }
                // 添加styleBean
                styleMap.put(styleId, styleBean);
            }
        }
        // 返回结果
        return styleMap;
    }

    /**
     * 获取表头数据
     * 
     * @param title 标题节点
     * @param params 入参map
     * @return List<RowBean>
     * @author andy
     * @throws Exception
     * @date 2017年1月18日
     */
    @SuppressWarnings("unchecked")
    private static List<RowBean> getRows(Element title, Map<String, Object> params) throws Exception {
        // 返回结果表头数据
        List<RowBean> rowList = new ArrayList<RowBean>();
        // 获取标题节点下的全部行节点
        List<Element> rows = title.elements("ROW");
        // 行数据
        RowBean rowBean = null;
        // 遍历行节点集合
        for (Element row : rows) {
            // 创建行数据
            rowBean = new RowBean();
            // 获取行号
            String rowNum = row.attributeValue("rownum");
            // 判断行号
            if(EmptyUtil.isNotEmpty(rowNum)){
            	// 设置行号
            	rowBean.setRowNum(Integer.parseInt(rowNum));
            }
            // 获取样式id
            String style = row.attributeValue("style");
            // 判断样式id
            if(EmptyUtil.isNotEmpty(style)){
            	// 设置样式id
                rowBean.setStyleId(style);
            }
            // 获取单元格数据
            List<CellBean> cellList = getCell(row, params);
            // 设置单元格数据
            rowBean.setChilds(cellList);
            // 添加行数据
            rowList.add(rowBean);
        }
        // 返回数据
        return rowList;
    }
    
    /**
     * 获取单元格数据
     * 
     * @param row 行节点
     * @param params 入参map
     * @return List<CellBean>
     * @author andy
     * @throws Exception
     * @date 2017年1月18日
     */
    @SuppressWarnings("unchecked")
    private static List<CellBean> getCell(Element row, Map<String, Object> params) throws Exception {
        // 返回结果单元格数据
        List<CellBean> cellList = new ArrayList<CellBean>();
        // 获取行节点下的全部列节点
        List<Element> cols = row.elements("COL");
        // 单元格数据
        CellBean cellBean = null;
        // 遍历列节点集合
        for (Element col : cols) {
            // 创建单元格数据
        	cellBean = new CellBean();
            // 获取列号
        	String colNum = col.attributeValue("colnum");
        	// 判断列号
        	if(EmptyUtil.isNotEmpty(colNum)){
        		// 设置列号
        		cellBean.setColNum(Integer.parseInt(colNum));
        	}
            // 获取行合并值
        	String rowspan = col.attributeValue("rowspan");
        	// 判断行合并值
        	if(EmptyUtil.isNotEmpty(rowspan)){
        		// 设置行合并值
        		cellBean.setRowspan(Integer.parseInt(rowspan));
        	}
            // 获取列合并值
            String colspan = col.attributeValue("colspan");
            // 判断列合并值
        	if(EmptyUtil.isNotEmpty(colspan)){
        		// 设置列合并值
        		cellBean.setColspan(Integer.parseInt(colspan));
        	}
            // 获取节点值
            String value = col.getTextTrim();
            if(EmptyUtil.isNotEmpty(value) && value.indexOf("@")!=-1){
            	// 获取入参map中key值集
				Set<String> keys = params.keySet();
				// 遍历key值集
				for (String key : keys) {
				    // 处理字符
					value = value.replaceAll("@" + key, params.get(key).toString());
				}
            }
            // 设置节点值
            cellBean.setValue(value);
            // 添加单元格数据
            cellList.add(cellBean);
        }
        return cellList;
    }

    /**
     * 获取填充数据区
     * 
     * @param sheet sheet节点
     * @param params
     * @return DataBean
     * @author andy
     * @throws Exception
     * @date 2017年1月18日
     */
    private static DataBean getData(Element sheet, Map<String, Object> params) throws Exception {
        // 返回结果填充数据区
        DataBean dataBean = new DataBean();
        // 获取填充数据区节点
        Element data = sheet.element("DATA");
        // 获取行起点
        String rowNum = data.attributeValue("rownum");
        // 判断行起点
        if(EmptyUtil.isNotEmpty(rowNum)){
        	// 设置行起点
            dataBean.setRowNum(Integer.parseInt(rowNum));
        }
        // 获取列起点
        String colNum = data.attributeValue("colnum");
        // 判断列起点
        if(EmptyUtil.isNotEmpty(colNum)){
        	// 设置列起点
            dataBean.setColNum(Integer.parseInt(colNum));
        }
        // 获取data填充方向
        String direction = data.attributeValue("direction");
        // 判断data填充方向
        if(EmptyUtil.isNotEmpty(direction)){
        	// 设置dataBean填充方向
            dataBean.setDirection(direction);
        }
        // 获取单元格样式id
        String styleId = data.attributeValue("style");
        // 判断单元格样式id
        if(EmptyUtil.isNotEmpty(styleId)){
        	// 设置单元格样式id
            dataBean.setStyleId(styleId);
        }
        // 返回结果
        return dataBean;
    }
}
