package converter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;

public class GenerateExcel
{
    HSSFRow row1;
    HSSFRow row2;
    HSSFWorkbook wb;
    HSSFSheet sheet;
    HSSFFont font;
    CellStyle style;
    GenerateExcelPojo pojo;
	
    public  void generate()
    {
    	 try
         {
    		pojo=new GenerateExcelPojo();
 	        System.out.println(pojo);
 	        pojo.setAgent("Sut2");
 	        pojo.setAmount("500");
 	        pojo.setBene_Address("Dwarka sector 7");
 	        pojo.setBene_ID("Bene ID");
 	        pojo.setBene_Name_1("Kapilraj");
 	        pojo.setBene_Phone("885588588");
 	        pojo.setCurrency("USD");
 	        pojo.setDate("12/1/2014");
 	        pojo.setSender_Name("Sumit Rawat");
 	        pojo.setInvoice_Number("457411256");     
 	        
 	       File file =new File("D:/eXCELfILE.xls");
 	 
 	    if(file.exists())
 	     {
 	    	System.out.println("File already exist");
 	    	wb=getWorkbook("D:/eXCELfILE.xls");
 	 		sheet=wb.getSheetAt(0);
 	 		int rowNum = sheet.getLastRowNum();
 	        System.out.println(rowNum);
 	 		row2=sheet.createRow(rowNum+1);
 	     }
 	    else
 	     {
 	    	System.out.println("File was created ");
 	    
 	    	wb = getNewWorkbook("D:/eXCELfILE.xls");
 	    	sheet = wb.createSheet();
 	    	font = wb.createFont();
	 		style=wb.createCellStyle();
	 		
	 		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
	 		font.setItalic(false);
	 		font.setColor((short)808080);
	 		style.setFont(font);

 	   		row1 = sheet.createRow(0);
 	   		//row1.setRowStyle(style);
 	   		
 	   		row2 = sheet.createRow(1);
 	   		getDefaultRow();
 	      }
 	      	getRow();

	 	    FileOutputStream out=new FileOutputStream("D:/eXCELfILE.xls");
			wb.write(out);
			out.close();
	        System.out.println("written successfully on disk.");
        }
        
        catch (Exception e) 
        {
            e.printStackTrace();
        }
    }
	private void getRow()
	{
		    Cell[] cell = getCellArray();
		    
		    for(int i=0;i<cell.length;i++)
		    {
		    	cell[i]= row2.createCell(i);
		    	
		    }

		    cell[0].setCellValue(pojo.getAgent());
		   	cell[1].setCellValue(pojo.getAmount());
		   	cell[2].setCellValue(pojo.getBene_Address());
		   	cell[3].setCellValue(pojo.getBene_ID());
		   	cell[4].setCellValue(pojo.getBene_Name_1());
		   	cell[5].setCellValue(pojo.getBene_Phone());
		   	cell[6].setCellValue(pojo.getCurrency());
		   	cell[7].setCellValue(pojo.getDate());
		   	cell[8].setCellValue(pojo.getInvoice_Number());
		   	cell[9].setCellValue(pojo.getSender_Name());
		   	cell[10].setCellValue(pojo.getNote());

	}
	private  void getDefaultRow()
	{
			Cell[] cell = getCellArray();
		    
		    for(int i=0;i<cell.length;i++)
		    {
		    	cell[i]= row2.createCell(i);
		    	
		    }
		
	        cell[0].setCellValue("Agent");
	       	cell[1].setCellValue("Amount");
	       	cell[2].setCellValue("Bene Address");
	       	cell[3].setCellValue("Bene ID");
	       	cell[4].setCellValue("Bene Name");
	       	cell[5].setCellValue("Bene Phone");
	       	cell[6].setCellValue("Currency");
	       	cell[7].setCellValue("Transaction Date");
	       	cell[8].setCellValue("Invoice Number");
	       	cell[9].setCellValue("Sender Name");
	       	cell[10].setCellValue("Note");

	       	for(int i = 0; i < row1.getLastCellNum(); i++){
	            row1.getCell(i).setCellStyle(style);
	        }
	}
	
	private static Cell[] getCellArray()
	{
		
		return new Cell[12];
	}
	
	private static HSSFWorkbook getWorkbook(String excelFilePath) throws IOException 
	{
		HSSFWorkbook workbook = null;
		
		if (excelFilePath.endsWith("xls")) 
			{
				FileInputStream io=new FileInputStream(excelFilePath);
				workbook = new HSSFWorkbook(io);
			}
		else
			{
				throw new IllegalArgumentException("Please give the Excel File");
			}
		
		return workbook;
	}
	
	private static HSSFWorkbook getNewWorkbook(String excelFilePath) throws IOException 
	{
		HSSFWorkbook workbook = null;

		if (excelFilePath.endsWith("xls")) 
		{
			workbook = new HSSFWorkbook();
		}
		else
		{
			throw new IllegalArgumentException("Please give the Excel File");
		}
		
		return workbook;
	}
	public static void main(String[] args)
	{
		GenerateExcel get=new GenerateExcel();
		get.generate();
		
	}
}
