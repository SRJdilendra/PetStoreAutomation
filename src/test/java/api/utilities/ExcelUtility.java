package api.utilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.text.StyleConstants;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtility {

	public FileInputStream FIS;
	public FileOutputStream FOS;
	public XSSFWorkbook workbook;
	public XSSFSheet sheet;
	public XSSFRow row;
	public XSSFCell cell;
	String path;
	
	public ExcelUtility(String path) {
		this.path=path;
	}
	public int getRowCount(String sheetName) throws IOException {
		
		FIS=new FileInputStream(path);
		workbook=new XSSFWorkbook(FIS);
		sheet=workbook.getSheet(sheetName);
		int rowcount=sheet.getLastRowNum();
		workbook.close();
		FIS.close();
		return rowcount;
		
	}
	public int getCellCount(String sheetName,int rownum) throws IOException {
		FIS=new FileInputStream(path);
		workbook=new XSSFWorkbook(FIS);
		sheet=workbook.getSheet(sheetName);
		int cellcount=row.getLastCellNum();
		workbook.close();
		FIS.close();
		return cellcount;
	}
	
	public String getCellData(String sheetName, int ronum, int colnum) throws IOException {
		FIS=new FileInputStream(path);
		workbook=new XSSFWorkbook(FIS);
		sheet=workbook.getSheet(sheetName);
		row=sheet.getRow(ronum);
		cell=row.getCell(colnum);
		
		DataFormatter DF=new DataFormatter();
		String Data;
		try {
			Data =DF.formatCellValue(cell); // Return the formated value of a cell as string regardless
			
		} catch (Exception e) {
			Data="";
		}
		workbook.close();
		FIS.close();
		return Data;		
	}
	/*
	public void setCellData(String sheetName, int rounum, int colnum, String data) throws IOException {
		File excelFile=new File(path);
		if (!excelFile.exists()) { // If file does not exist then create file
			
			workbook=new XSSFWorkbook();
			FOS=new FileOutputStream(path);
			workbook.write(FOS);
		}
		FIS=new FileInputStream(path);
		workbook=new XSSFWorkbook(FIS);
		
		if(workbook.getSheetIndex(sheetName)==-1) // If sheet is not exist then create new sheet
			workbook.createSheet(sheetName);
		sheet=workbook.getSheet(sheetName);
		
		if(sheet.getRow(rounum)==null) // If row not exists then create row
			sheet.createRow(rounum);
		row=sheet.getRow(rounum);
		
		cell=row.createCell(colnum);
		cell.setCellValue(data);
		FOS=new FileOutputStream(path);
		workbook.write(FOS);
		workbook.close();
		FIS.close();
		FOS.close();
	}
	public void fillGreenColour(String sheetName, int rownum, int colnum) {
		FIS=new FileInputStream(path);
		workbook=new XSSFWorkbook(FIS);
		sheet=workbook.getSheet(sheetName);
		row=sheet.getRow(rownum);
		cell=row.getCell(colnum);
		Styles=workbook.createCellStyle();
		styles.  // Not completed fully some code is remaining i have taken screenshot, later i will completec
	}
	*/
}
