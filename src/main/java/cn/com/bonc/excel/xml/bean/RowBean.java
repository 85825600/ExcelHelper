/*****************************************************************************
 * 东方国信Excel助手[ExcelHelper]
 *----------------------------------------------------------------------------
 * cn.com.bonc.excel.xml.bean.RowBean.java
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

import java.util.List;

/**
 * 行数据集Bean
 * cn.com.bonc.excel.xml.bean.RowBean.java
 * 
 * @author andy
 * @date 2017年1月18日
 *
 * @since 0.0.1
 */
public class RowBean extends BaseBean {
    /**
     * 行下列集合
     */
    private List<CellBean> childs;

	/**
	 * @return the childs
	 */
	public List<CellBean> getChilds() {
		return childs;
	}

	/**
	 * @param childs the childs to set
	 */
	public void setChilds(List<CellBean> childs) {
		this.childs = childs;
	}
}
