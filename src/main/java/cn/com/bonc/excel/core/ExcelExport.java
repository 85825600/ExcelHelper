/*****************************************************************************
 * 东方国信Excel助手[ExcelHelper]
 *----------------------------------------------------------------------------
 * cn.com.bonc.excel.core.ExcelExport.java
 *
 * @author andy
 * @date 2017年1月17日
 * @version 0.0.1
 * @since 0.0.1
 *----------------------------------------------------------------------------
 * (C) 北京东方国信科技股份有限公司
 *     Business-intelligence Of Oriental Nations Corporation Ltd. 2016
 *****************************************************************************/
package cn.com.bonc.excel.core;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.ss.usermodel.CellStyle;

import cn.com.bonc.excel.xml.XmlReader;
import cn.com.bonc.excel.xml.bean.CellBean;
import cn.com.bonc.excel.xml.bean.RowBean;
import cn.com.bonc.excel.xml.bean.SheetBean;
import cn.com.bonc.excel.xml.bean.StyleBean;

/**
 * Excel导出类
 * cn.com.bonc.excel.core.ExcelExport.java
 * 
 * @author andy
 * @date 2017年1月17日
 *
 * @since 0.0.1
 */
public class ExcelExport {
	/**
	 * 导出excel方法
	 * 
	 * @param fielName xml模板文件名
	 * @param paramsMap 传入参数map
	 * @param datas 数据区数据
	 * @return HSSFWorkbook
	 * @author andy
	 * @throws Exception
	 * @date 2017年1月18日
	 */
    @SuppressWarnings("unchecked")
	public static HSSFWorkbook exportExcel(String fileName, Map<String, Object> paramsMap, List<List<Object>>... datas) throws Exception {
        InputStream inputStream = ExcelExport.class.getClassLoader().getResourceAsStream(fileName);
        return exportExcel(inputStream, paramsMap, datas);
    }

    /**
     * 导出excel方法
     * 
     * @param inputStream xml文件输入流
     * @param paramsMap 传入参数map
     * @param datas 数据区数据
     * @return HSSFWorkbook
     * @author andy
     * @throws Exception
     * @date 2017年1月18日
     */
    @SuppressWarnings("unchecked")
	public static HSSFWorkbook exportExcel(InputStream inputStream, Map<String, Object> paramsMap, List<List<Object>>... datas) throws Exception {
    	// 创建excel对象
    	HSSFWorkbook excel = new HSSFWorkbook();
		// 解析xml获取excel模板数据
		List<SheetBean> sheets = XmlReader.readXml(inputStream, paramsMap);
		// sheet页
		HSSFSheet sheet = null;
		// 遍历excel模板数据
		for (int i=0; i<sheets.size(); i++) {
			
			// sheet模板数据
			SheetBean sheetBean = sheets.get(i);
		    // 创建sheet页
		    sheet = excel.createSheet(sheetBean.getSheetName());
		    // 填充sheet标题数据
		    setSheetTitle(excel, sheet, sheetBean);
		    // 数据区数据
		 	List<List<Object>> data = new ArrayList<List<Object>>();
		 	if(i<datas.length){
		 		data = datas[i];
		 	}
		    // 填充sheet数据区数据
		    setSheetData(excel, sheet, sheetBean, data);
		    // 调整数据区单元格样式
		    setSheetDataStyle(excel, sheet, sheetBean, data);
		}
		// 返回excel
		return excel;
    }

    /**
     * 填充sheet标题数据
     * 
     * @param excel excel对象
     * @param sheet excel的sheet页
     * @param sheetBean sheet模板数据
     * @author andy
     * @throws Exception
     * @date 2017年1月21日
     */
    private static void setSheetTitle(HSSFWorkbook excel, HSSFSheet sheet, SheetBean sheetBean) throws Exception {
        // 获取单元格样式map
        Map<String, StyleBean> styles = sheetBean.getStyles();
        // 获取xml中标题行数据
        List<RowBean> rows = sheetBean.getTitle();
        // 遍历标题行数据
        for (RowBean row : rows) {
        	// 获取行号
        	Integer rowNum = row.getRowNum();
			// 获取单元格样式id
			String styleId = row.getStyleId();
			// 创建单元格样式
			CellStyle cellStyle = createCellStyle(excel, styles, styleId);
			// 获取row下子节点
			List<CellBean> childs = row.getChilds();
			// 遍历row下子节点
			for (CellBean child : childs) {
			    // 获取列号
			    Integer colNum = child.getColNum();
			    // 获取行合并值
			    Integer rowspan = child.getRowspan();
			    // 获取列合并值
			    Integer colspan = child.getColspan();
			    // 获取节点值
			    Object value = child.getValue();
			    // 合并单元格
			    mergeCell(sheet, rowNum, colNum, rowspan, colspan);
			    // 填充单元格内容
			    fillCell(sheet, rowNum, colNum, rowspan, colspan, cellStyle, value.toString());
			}
        }
    }

    /**
     * 填充sheet数据区数据
     * 
     * @param excel excel对象
     * @param sheet excel的sheet页
     * @param sheetBean sheet模板数据
     * @param datas 数据区数据
     * @author andy
     * @throws Exception
     * @date 2017年1月21日
     */
    private static void setSheetData(HSSFWorkbook excel, HSSFSheet sheet, SheetBean sheetBean, List<List<Object>> datas) throws Exception {
    	// 获取行起点
    	Integer rowNum = sheetBean.getData().getRowNum();
    	// 获取列起点
    	Integer colNum = sheetBean.getData().getColNum();
    	// 获取数据填充方向
		String direction = sheetBean.getData().getDirection();
		// 获取单元格样式id
		String styleId = sheetBean.getData().getStyleId();
		// 获取单元格样式map
		Map<String, StyleBean> styles = sheetBean.getStyles();
		// 创建单元格样式
		CellStyle cellStyle = createCellStyle(excel, styles, styleId);
		// 判断填充方向
		if("row".equalsIgnoreCase(direction)){
			// 列号
			int c = colNum;
			// 遍历数据数据区数据
			for(List<Object> col : datas){
				// 行号
				int r = rowNum;
				// 遍历列单元格
				for(Object cell : col){
					// 填充单元格
					fillCell(sheet, r, c, 1, 1, cellStyle, cell.toString());
					// 行号+1
					r++;
				}
				// 列号+1
				c++;
			}
		} else {
			// 行号
			int r = rowNum;
			// 遍历数据数据区数据
			for(List<Object> row : datas){
				// 列号
				int c = colNum;
				// 遍历行单元格
				for(Object cell : row){
					// 填充单元格
					fillCell(sheet, r, c, 1, 1, cellStyle, cell.toString());
					// 列号+1
					c++;
				}
				// 行号+1
				r++;
			}
		}
    }

    /**
     * 创建单元格样式
     * 
     * @param excel excel对象
     * @param styles 单元格样式map
     * @param styleId 单元格样式id
     * @param CellStyle
     * @author andy
     * @throws Exception
     * @date 2017年1月21日
     */
    private static CellStyle createCellStyle(HSSFWorkbook excel, Map<String, StyleBean> styles, String styleId) throws Exception {
        // 创建单元格样式
        CellStyle style = excel.createCellStyle();
        // 创建字体
        HSSFFont font = excel.createFont();
        // style模板数据
        StyleBean styleBean = styles.get(styleId);
        // 设置字体大小
        font.setFontHeightInPoints(styleBean.getFontSize());
        // 设置字体颜色
        font.setColor(styleBean.getFontColor());
        // 设置字体加粗值
        font.setBoldweight(styleBean.getBoldWeight());
        // 设置字体
        style.setFont(font);
        // 设置单元格对齐方式
        style.setAlignment(styleBean.getAlign());
        // 设置垂直对齐方式(固定为居中)
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        // 设置单元格背景色
        style.setFillForegroundColor(styleBean.getBackgroundColor());
        // 填充方式
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        // 判断有无边框
        if (styleBean.getBorder()) {
            // 下边框
            style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            // 左边框
            style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            // 上边框
            style.setBorderTop(HSSFCellStyle.BORDER_THIN);
            // 右边框
            style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        }
        // 返回样式
        return style;
    }
    
    /**
     * 判断是否需要合并单元格
     * 
     * @param span 合并值
     * @return Boolean
     * @author andy
     * @date 2017年1月21日
     */
    private static Boolean isNeedMerge(int span) {
        // 判断合并值
        if (span > 1) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * 合并单元格
     * 
     * @param sheet excel的sheet页
     * @param rowNum 行号
     * @param colNum 列号
     * @param rowspan 行合并值
     * @param colspan 列合并值
     * @author andy
     * @throws Exception
     * @date 2017年1月21日
     */
    @SuppressWarnings("deprecation")
    private static void mergeCell(HSSFSheet sheet, int rowNum, int colNum, int rowspan, int colspan) throws Exception {
        // 判断行是否需要合并单元格
        if (isNeedMerge(rowspan)) {
            // 判断列是否需要合并单元格
            if (isNeedMerge(colspan)) {
                // 行列都需要合并的情况
                sheet.addMergedRegion(new Region(rowNum, (short)colNum, rowNum+rowspan-1, (short)(colNum+colspan-1)));
            } else {
                // 行需要合并的情况
                sheet.addMergedRegion(new Region(rowNum, (short)colNum, rowNum+rowspan-1, (short)colNum));
            }
        } else {
            if (isNeedMerge(colspan)) {
                // 列需要合并的情况
                sheet.addMergedRegion(new Region(rowNum, (short)colNum, rowNum, (short)(colNum+colspan-1)));
            }
        }
        sheet.setColumnWidth(colNum, 20*256);
    }

    /**
     * 填充单元格
     * 
     * @param sheet excel的sheet页
     * @param rowNum 行号
     * @param colNum 列号
     * @param rowspan 行合并值
     * @param colspan 列合并值
     * @param style 单元格样式
     * @param value 单元格内容
     * @author andy
     * @throws Exception
     * @date 2017年1月21日
     */
    private static void fillCell(HSSFSheet sheet, int rowNum, int colNum, int rowspan, int colspan, CellStyle style, String value) throws Exception {
        // 设定列宽
        sheet.setColumnWidth(0, 3000);
        for (int r = rowNum; r < rowNum+rowspan; r++) {
            // 获取行
            HSSFRow row = sheet.getRow(r);
            // 判断存不存在行
            if (row == null) {
                row = sheet.createRow(r);
            }
            for (int c = colNum; c < colNum+colspan; c++) {
                // 获取单元格
                HSSFCell cell = row.getCell(c);
                // 判断有无单元格
                if (cell == null) {
                    cell = row.createCell(c);
                }
                // 设置单元格样式
                cell.setCellStyle(style);
                // 设置单元格内容
                cell.setCellValue(value);
            }
        }
    }

    /**
     * 调整数据区单元格样式
     * 
     * @param excel excel对象
     * @param sheet excel的sheet页
     * @param rowLength 行长度
     * @author andy
     * @throws Exception
     * @date 2017年1月21日
     */
    private static void setSheetDataStyle(HSSFWorkbook excel, HSSFSheet sheet, SheetBean sheetBean, List<List<Object>> datas) throws Exception {
        // 获取xml中数据的行起点
        Integer rowNum = sheetBean.getData().getRowNum();
        // 获取xml中数据的列起点
        Integer colNum = sheetBean.getData().getColNum();
        // 获取填充方向
        String direction = sheetBean.getData().getDirection();
        // 获取数据区域样式id
        String styleId = sheetBean.getData().getStyleId();
        // 获取单元格样式map
        Map<String, StyleBean> styles = sheetBean.getStyles();
        // 创建单元格样式
        CellStyle cellStyle = createCellStyle(excel, styles, styleId);
        // 数据记录数
        int recordNum = 0;
        // 记录字段数
        int fieldNum = 0;
        // 判断数据区数据是否为空
        if(datas!=null && !datas.isEmpty()){
        	// 获取数据记录数
        	recordNum = datas.size();
        	// 获取首个记录
        	List<Object> record = datas.get(0);
        	// 判断记录是否为空
        	if(record != null){
        		// 获取记录字段数
        		fieldNum = record.size();
        	}
        }
        if(recordNum*fieldNum != 0){
        	// 判断填充方向
            if ("row".equalsIgnoreCase(direction)) {
                // 行终点
                int endRow = rowNum+fieldNum-1;
                // 列终点
                int endCol = colNum+recordNum-1;
                // 设置单元格样式
                fillStyle(sheet, rowNum, colNum, endRow, endCol, cellStyle);
            } else {
            	// 行终点
                int endRow = rowNum+recordNum-1;
                // 列终点
                int endCol = colNum+fieldNum-1;
                // 设置单元格样式
                fillStyle(sheet, rowNum, colNum, endRow, endCol, cellStyle);
            }
        }
        
    }

    /**
     * 补充单元格样式
     * 
     * @param sheet excel的sheet页
     * @param beginRow 行起点
     * @param beginCol 列起点
     * @param endRow 行终点
     * @param endCol 列终点
     * @param style 单元格样式
     * @author andy
     * @throws Exception
     * @date 2017年1月21日
     */
    private static void fillStyle(HSSFSheet sheet, int beginRow, int beginCol, int endRow, int endCol, CellStyle style) throws Exception {
        // 循环填充每一个单元格样式
        for (int r = beginRow; r <= endRow; r++) {
            // 获取行
            HSSFRow row = sheet.getRow(r);
            // 判断存不存在行
            if (row == null) {
                row = sheet.createRow(r);
            }
            for (int c = beginCol; c < endCol; c++) {
                // 获取单元格
                HSSFCell cell = row.getCell(c);
                // 判断有无单元格
                if (cell == null) {
                    cell = row.createCell(c);
                }
                // 设置单元格样式
                cell.setCellStyle(style);
            }
        }
    }
}
