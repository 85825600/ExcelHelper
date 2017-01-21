/*****************************************************************************
 * 东方国信Excel助手[ExcelHelper]
 *----------------------------------------------------------------------------
 * cn.com.bonc.excel.ExcelHelperTest.java
 *
 * @author andy
 * @date 2017年1月21日
 * @version 0.0.1
 * @since 0.0.1
 *----------------------------------------------------------------------------
 * (C) 北京东方国信科技股份有限公司
 *     Business-intelligence Of Oriental Nations Corporation Ltd. 2016
 *****************************************************************************/
package cn.com.bonc.excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;

/**
 * Excel助手测试类
 * cn.com.bonc.excel.ExcelHelperTest.java
 * 
 * @author andy
 * @date 2017年1月21日
 *
 * @since 0.0.1
 */
public class ExcelHelperTest {
	@Test
	public void exportExcel(){
		FileOutputStream fos = null;
        try {
            long start = System.currentTimeMillis();
            // 传入参数map
            Map<String, Object> params = new HashMap<String, Object>();
            params.put("date", "2017年1月");
            params.put("city", "辽宁省");
            params.put("target", "某指标");
            // 数据区数据
            List<List<Object>> datas = new ArrayList<List<Object>>();
            HSSFWorkbook excel = ExcelHelper.exportExcel("template/demo.xml", params, datas);
            fos = new FileOutputStream("/Users/andy/Downloads/test.xls");
            excel.write(fos);
            fos.flush();
            long end = System.currentTimeMillis();
            System.out.println("导出excel所用时间：" + (end-start) + "ms");
        } catch (Exception e) {
			e.printStackTrace();
        } finally {
        	try {
				fos.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
        }
	}
	
	public static void main(String[] args) {
		new ExcelHelperTest().exportExcel();
	}
}
