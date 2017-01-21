/*****************************************************************************
 * 东方国信Excel助手[ExcelHelper]
 *----------------------------------------------------------------------------
 * cn.com.bonc.excel.xml.bean.SheetBean.java
 *
 * @author andy
 * @date 2017年1月17日
 * @version 0.0.1
 * @since 0.0.1
 *----------------------------------------------------------------------------
 * (C) 北京东方国信科技股份有限公司
 *     Business-intelligence Of Oriental Nations Corporation Ltd. 2016
 *****************************************************************************/
package cn.com.bonc.excel.xml.bean;

import java.util.List;
import java.util.Map;

/**
 * Sheet页数据集Bean
 * cn.com.bonc.excel.xml.bean.SheetBean.java
 * 
 * @author andy
 * @date 2017年1月17日
 *
 * @since 0.0.1
 */
public class SheetBean {
	/**
     * sheet页名称
     */
    private String sheetName = "";
    
    /**
     * 单元格样式map
     */
    private Map<String, StyleBean> styles;

    /**
     * 表头数据
     */
    private List<RowBean> title;

    /**
     * 数据区bean
     */
    private DataBean data;

	/**
	 * @return the sheetName
	 */
	public String getSheetName() {
		return sheetName;
	}

	/**
	 * @param sheetName the sheetName to set
	 */
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	/**
	 * @return the style
	 */
	public Map<String, StyleBean> getStyles() {
		return styles;
	}

	/**
	 * @param style the style to set
	 */
	public void setStyles(Map<String, StyleBean> styles) {
		this.styles = styles;
	}
	
	/**
	 * @return the title
	 */
	public List<RowBean> getTitle() {
		return title;
	}

	/**
	 * @param title the title to set
	 */
	public void setTitle(List<RowBean> title) {
		this.title = title;
	}

	/**
	 * @return the data
	 */
	public DataBean getData() {
		return data;
	}

	/**
	 * @param data the data to set
	 */
	public void setData(DataBean data) {
		this.data = data;
	}
}
