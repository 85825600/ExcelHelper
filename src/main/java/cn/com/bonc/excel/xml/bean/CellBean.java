/*****************************************************************************
 * 东方国信Excel助手[ExcelHelper]
 *----------------------------------------------------------------------------
 * cn.com.bonc.excel.xml.bean.CellBean.java
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
 * 单元格Bean
 * cn.com.bonc.excel.xml.bean.CellBean.java
 * 
 * @author andy
 * @date 2017年1月18日
 *
 * @since 0.0.1
 */
public class CellBean extends BaseBean {
	/**
     * 行合并值
     */
    private Integer rowspan = 1;

    /**
     * 列合并值
     */
    private Integer colspan = 1;
    
    /**
     * 节点值
     */
    private String value = "";

	/**
	 * @return the rowspan
	 */
	public Integer getRowspan() {
		return rowspan;
	}

	/**
	 * @param rowspan the rowspan to set
	 */
	public void setRowspan(Integer rowspan) {
		this.rowspan = rowspan;
	}

	/**
	 * @return the colspan
	 */
	public Integer getColspan() {
		return colspan;
	}

	/**
	 * @param colspan the colspan to set
	 */
	public void setColspan(Integer colspan) {
		this.colspan = colspan;
	}

	/**
	 * @return the value
	 */
	public String getValue() {
		return value;
	}

	/**
	 * @param value the value to set
	 */
	public void setValue(String value) {
		this.value = value;
	}
}
