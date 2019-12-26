package Utilities;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.text.Style;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtil {
	   Workbook wb;
	   //constructor for accessing excelpath
	   public ExcelFileUtil() throws Throwable{
		   FileInputStream fi=new FileInputStream(System.getProperty("user.dir")+"\\TestInput\\InputSheet.xlsx");
		   wb=WorkbookFactory.create(fi);
	   }
	   //count no of rows in a sheet
	   public int rowCount(String sheetname)
	   {
		   return wb.getSheet(sheetname).getLastRowNum();
		   
	   }
	   //count no of column in a sheet
	   public int colCount(String sheetname)
	   {
		   return wb.getSheet(sheetname).getRow(0).getLastCellNum();
	   }
	   //get data from cell
	   public String getCellData(String sheetname,int row,int column)
	   {
		 String data="";
		 if(wb.getSheet(sheetname).getRow(row).getCell(column).getCellType()==Cell.CELL_TYPE_NUMERIC)
		 {
			 int celldata=(int)wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();
			 data=String.valueOf(celldata);
		 }
		 else
		 {
			 data=wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();
		 }
			 return data;
			
	   }
	   //writing results into cell
	   public void setCellData(String sheetname,int row,int column,String status) throws Throwable{
		   Sheet ws=wb.getSheet(sheetname);
		   Row rownum=ws.getRow(row);
		   Cell cell=rownum.createCell(column);
		   cell.setCellValue(status);
		   if(status.equalsIgnoreCase("pass")){
			   //create a cell style
			   CellStyle style=wb.createCellStyle();
			   //create a font
			   Font font=wb.createFont();
			   //apply color to nthe text
			   font.setColor(IndexedColors.GREEN.getIndex());
			   //apply bold to the text
			   font.setBold(true);
			   //set font
			   style.setFont(font);
			   //set cell type
			   rownum.getCell(column).setCellStyle(style);
		   }
		   FileOutputStream fo=new FileOutputStream(System.getProperty("user.dir")+"\\TestOutput\\Hybrid.xlsx");
		   wb.write(fo);
		   fo.close();
		   
			   
			   
		   }
	   
	   
	   
		   
		 

	   
	   
	   
	   
	   
	   
	   
	   
	   
	   
	   
	   
	   
	   
	   


}


