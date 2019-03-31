package Poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sample 
{
	public static void main(String args[]) throws IOException
	{
		FileInputStream file = new FileInputStream("C:\\Users\\SATHIYA\\Desktop\\sample.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		int noOfRows = sheet.getPhysicalNumberOfRows();
		String[] Labelval = new String[noOfRows];
		int noOfColumns = sheet.getRow(0).getLastCellNum();
		String[] Headers = new String[noOfColumns];
		String val = null;
		for(int i=0;i<noOfRows;i++)
		{
			Labelval[i] = sheet.getRow(i).getCell(0).getStringCellValue();
			if(Labelval[i].equals("sathya"))
			{
				for(int j=0;j<noOfColumns;j++) 
					Headers[j] = sheet.getRow(0).getCell(j).getStringCellValue();
				for(int a=0;a<noOfColumns;a++)
				{
					if(Headers[a].equals("ice"))
					{
						val = sheet.getRow(i).getCell(a).getStringCellValue();
						System.out.println(sheet.getRow(i).getCell(a).getStringCellValue());
					}
				}
			}
		}
	}
}
