package d_C_P;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_Utils {
	XSSFWorkbook wb;
	XSSFSheet s;
	Excel_Utils(String excelPath,String sheetName) throws Exception
	{
		FileInputStream fis=new FileInputStream(excelPath);
		wb=new XSSFWorkbook(fis);
		s=wb.getSheet(sheetName);
	}
	public void setCellData(int rowindex,int colindex,String excelPath,String result) throws IOException
	{
		s.getRow(rowindex).createCell(colindex).setCellValue(result);
		FileOutputStream fos=new FileOutputStream(excelPath);
		wb.write(fos);
	}
	public String getCellData(int rowindex,int colindex)
	{
		return s.getRow(rowindex).getCell(colindex).getStringCellValue();
	}
	public int getRowCount()
	{
		return s.getLastRowNum();
	}



}
