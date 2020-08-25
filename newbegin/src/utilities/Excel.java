package utilities;

import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
	
	private XSSFSheet excelSheet;
	
	public Excel (String path,String sheetname)
	{
		try
		{
			FileInputStream ExcelFile = new FileInputStream(path);
			XSSFWorkbook excelWorkbook = new XSSFWorkbook (ExcelFile);
			excelSheet = excelWorkbook.getSheet(sheetname);
		}
		catch(Exception e) {
			e.printStackTrace();
			
			 }
		public String getCellDatastring(int RowNum,int ColNum) {
			try {
				String CellData = excelSheet.getRow(RowNum).getCell(ColNum).getStringCellValue();
				return;
			}catch(Exception e) {
				return;
			}
			
		}
	}
	

}
