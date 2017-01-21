/*****************************************************************************
 * 东方国信Excel助手[ExcelHelper]
 *----------------------------------------------------------------------------
 * cn.com.bonc.excel.xml.bean.BaseBean.java
 *
 * @author andy
 * @date 2017年1月18日
 * @version 0.0.1
 * @since 0.0.1
 *----------------------------------------------------------------------------
 * (C) 北京东方国信科技股份有限公司
 *     Business-intelligence Of Oriental Nations Corporation Ltd. 2016
 *****************************************************************************/
package cn.com.bonc.excel.xml.bean;

/**
 * 基础Bean
 * cn.com.bonc.excel.xml.bean.BaseBean.java
 * 
 * @author andy
 * @date 2017年1月18日
 *
 * @since 0.0.1
 */
public class BaseBean {
	/**
     * 行号(从0开始)
     */
    private Integer rowNum;

    /**
     * 列号(从0开始)
     */
    private Integer colNum;

    /**
     * 样式id
     */
    private String styleId;

	/**
	 * @return the rowNum
	 */
	public Integer getRowNum() {
		return rowNum;
	}

	/**
	 * @param rowNum the rowNum to set
	 */
	public void setRowNum(Integer rowNum) {
		this.rowNum = rowNum;
	}

	/**
	 * @return the colNum
	 */
	public Integer getColNum() {
		return colNum;
	}

	/**
	 * @param colNum the colNum to set
	 */
	public void setColNum(Integer colNum) {
		this.colNum = colNum;
	}

	/**
	 * @return the styleId
	 */
	public String getStyleId() {
		return styleId;
	}

	/**
	 * @param styleId the styleId to set
	 */
	public void setStyleId(String styleId) {
		this.styleId = styleId;
	}
}
