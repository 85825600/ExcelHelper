/*****************************************************************************
 * 东方国信Excel助手[ExcelHelper]
 *----------------------------------------------------------------------------
 * cn.com.bonc.excel.ExcelHelper.java
 *
 * @author andy
 * @date 2017年1月17日
 * @version 0.0.1
 * @since 0.0.1
 *----------------------------------------------------------------------------
 * (C) 北京东方国信科技股份有限公司
 *     Business-intelligence Of Oriental Nations Corporation Ltd. 2016
 *****************************************************************************/
package cn.com.bonc.excel;

import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import cn.com.bonc.excel.core.ExcelExport;

/**
 * Excel助手
 * cn.com.bonc.excel.ExcelHelper.java
 * 
 * @author andy
 * @date 2017年1月17日
 *
 * @since 0.0.1
 */
public class ExcelHelper {
	/**
	 * 导出excel方法
	 * 
	 * @param fielName xml模板文件名
	 * @param datas 数据区数据
	 * @return HSSFWorkbook
	 * @author andy
	 * @date 2017年1月21日
	 */
    public static HSSFWorkbook exportExcel(String fileName, List<List<Object>> datas) {
    	Map<String, Object> params = new HashMap<String, Object>();
        return exportExcel(fileName, params, datas);
    }
	
	/**
	 * 导出excel方法
	 * 
	 * @param fielName xml模板文件名
	 * @param params 传入参数map
	 * @param datas 数据区数据
	 * @return HSSFWorkbook
	 * @author andy
	 * @date 2017年1月21日
	 */
    public static HSSFWorkbook exportExcel(String fileName, Map<String, Object> params, List<List<Object>> datas) {
        InputStream inputStream = ExcelExport.class.getClassLoader().getResourceAsStream(fileName);
        return exportExcel(inputStream, params, datas);
    }
    
    /**
	 * 导出excel方法
	 * 
	 * @param inputStream xml文件输入流
	 * @param datas 数据区数据
	 * @return HSSFWorkbook
	 * @author andy
	 * @date 2017年1月21日
	 */
    public static HSSFWorkbook exportExcel(InputStream inputStream, List<List<Object>> datas) {
		Map<String, Object> params = new HashMap<String, Object>();
        return exportExcel(inputStream, params, datas);
    }
    
    /**
	 * 导出excel方法
	 * 
	 * @param inputStream xml文件输入流
	 * @param params 传入参数map
	 * @param datas 数据区数据
	 * @return HSSFWorkbook
	 * @author andy
	 * @date 2017年1月21日
	 */
    @SuppressWarnings("unchecked")
	public static HSSFWorkbook exportExcel(InputStream inputStream, Map<String, Object> params, List<List<Object>> datas) {
    	HSSFWorkbook excel = null;
		try {
			excel = ExcelExport.exportExcel(inputStream, params, datas);
		} catch (Exception e) {
			e.printStackTrace();
		}
        return excel;
    }
}
