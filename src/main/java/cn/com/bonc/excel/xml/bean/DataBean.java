/*****************************************************************************
 * 东方国信Excel助手[ExcelHelper]
 *----------------------------------------------------------------------------
 * cn.com.bonc.excel.xml.bean.DataBean.java
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
 * 数据集Bean
 * cn.com.bonc.excel.xml.bean.DataBean.java
 * 
 * @author andy
 * @date 2017年1月18日
 *
 * @since 0.0.1
 */
public class DataBean extends BaseBean {
    /**
     * 填充方向
     */
    private String direction = "col";

	/**
	 * @return the direction
	 */
	public String getDirection() {
		return direction;
	}

	/**
	 * @param direction the direction to set
	 */
	public void setDirection(String direction) {
		this.direction = direction;
	}
}
