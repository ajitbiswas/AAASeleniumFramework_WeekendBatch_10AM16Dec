package generic;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Lib implements IAutoConstant {
	
	public static FileInputStream fis;
	public static Workbook wb;
	public static String getPropertyValue(String key){
		String propertyValue="";
		try {
			fis = new FileInputStream(CONFIG_PATH);
			Properties prop = new Properties();
			prop.load(fis);
			propertyValue = prop.getProperty(key);
		} catch (Exception e) {
		}
		return propertyValue;
	}
	public static String getCellValue(String sheetname, int row, int column){
		String cellValue="";
		try {
			fis = new FileInputStream(EXCEL_PATH);
			wb = WorkbookFactory.create(fis);
			cellValue = wb.getSheet(sheetname).getRow(row).getCell(column).toString();
		} catch (Exception e) {
		}
		return cellValue;
	}
	public static int getRowCount(String sheetname){
		int rc = 0;
		try {
			fis = new FileInputStream(EXCEL_PATH);
			wb = WorkbookFactory.create(fis);
			rc = wb.getSheet(sheetname).getLastRowNum();
		} catch (Exception e) {
		}
		return rc;
	}
}
