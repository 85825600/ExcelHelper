/*****************************************************************************
 * 东方国信Excel助手[ExcelHelper]
 *----------------------------------------------------------------------------
 * cn.com.bonc.excel.xml.bean.StyleBean.java
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
 * 单元格样式Bean
 * cn.com.bonc.excel.xml.bean.StyleBean.java
 * 
 * @author andy
 * @date 2017年1月18日
 *
 * @since 0.0.1
 */
public class StyleBean {
    /**
     * 背景颜色
     */
    private Short backgroundColor;
    
    /**
     * 字体大小
     */
    private Short fontSize;

    /**
     * 字体颜色
     */
    private Short fontColor;

    /**
     * 字体加粗
     */
    private Short boldWeight;

    /**
     * 对齐方式
     */
    private Short align;

    /**
     * 是否有边框
     */
    private Boolean border;

	/**
	 * @return the backgroundColor
	 */
	public Short getBackgroundColor() {
		return backgroundColor;
	}

	/**
	 * @param backgroundColor the backgroundColor to set
	 */
	public void setBackgroundColor(Short backgroundColor) {
		this.backgroundColor = backgroundColor;
	}

	/**
	 * @return the fontSize
	 */
	public Short getFontSize() {
		return fontSize;
	}

	/**
	 * @param fontSize the fontSize to set
	 */
	public void setFontSize(Short fontSize) {
		this.fontSize = fontSize;
	}

	/**
	 * @return the fontColor
	 */
	public Short getFontColor() {
		return fontColor;
	}

	/**
	 * @param fontColor the fontColor to set
	 */
	public void setFontColor(Short fontColor) {
		this.fontColor = fontColor;
	}

	/**
	 * @return the boldWeight
	 */
	public Short getBoldWeight() {
		return boldWeight;
	}

	/**
	 * @param boldWeight the boldWeight to set
	 */
	public void setBoldWeight(Short boldWeight) {
		this.boldWeight = boldWeight;
	}

	/**
	 * @return the align
	 */
	public Short getAlign() {
		return align;
	}

	/**
	 * @param align the align to set
	 */
	public void setAlign(Short align) {
		this.align = align;
	}

	/**
	 * @return the border
	 */
	public Boolean getBorder() {
		return border;
	}

	/**
	 * @param border the border to set
	 */
	public void setBorder(Boolean border) {
		this.border = border;
	}
}
