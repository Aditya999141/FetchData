package ReadExcelOperation;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;



public class ExcelRead
{

	public static void main(String[] args) throws IOException 
	{
		String excelFilePath=".\\Data\\Orders Details.xlsx"; // get the location of file
		FileInputStream inputstream = new FileInputStream (excelFilePath); // refer that file using FileInputStream
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputstream);// get the workbook
		XSSFSheet sheet= workbook.getSheetAt(0);//XSSFSheet sheet= workbook.getSheet("Sheet1");
		
		//Using for loop
		
		int rows = sheet.getLastRowNum(); // get the last row from sheet
		int columns=sheet.getRow(1).getLastCellNum();// get the last cell from row
		
		
		
		for(int r=0;r<=rows;r++) // outer loop for row
		{
			XSSFRow row=sheet.getRow(r); 
			for(int c=0;c<columns;c++) // inner loop for cell
			{
				XSSFCell cell=row.getCell(c); //Getting the particular cell from row
				
				

				
				switch(cell.getCellType())
				
				{
				
				case STRING: System.out.print(cell.getStringCellValue());break;
				
				case NUMERIC: System.out.print(cell.getNumericCellValue());break;
				
				default:System.out.println("*");break;
					
				
							
				}
				
				System.out.print("|");
			}
			
		}
		System.out.print(" | ");
		
		
	}

}
