package Poi;

import java.io.FileOutputStream;
import java.io.IOException;


import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateXL 
{
	public static void main(String[] args) throws IOException {
		Workbook wb=new XSSFWorkbook();
		FileOutputStream f=new FileOutputStream("C:\\Users\\SATHIYA\\Desktop\\CreateXLfile.xlsx");
		Sheet sh= wb.createSheet("Sheet1");
		Row r1=sh.createRow(0);
		Row r2=sh.createRow(1);
		Row r3=sh.createRow(2);
		r1.createCell(0).setCellValue("S.No");
		r1.createCell(1).setCellValue("Name");
		r1.createCell(2).setCellValue("Ph.No");
		r1.createCell(3).setCellValue("Email ID");
		r1.createCell(4).setCellValue("Place");
		
		r2.createCell(0).setCellValue("1");
		r2.createCell(1).setCellValue("Manoharan");
		r2.createCell(2).setCellValue("9786409698");
		r2.createCell(3).setCellValue("mano.hsm@gmail.com");
		r2.createCell(4).setCellValue("Trichy");
		
		r3.createCell(0).setCellValue("2");
		r3.createCell(1).setCellValue("Arun");
		r3.createCell(2).setCellValue("9876543210");
		r3.createCell(3).setCellValue("hi.hello@gmail.com");
		r3.createCell(4).setCellValue("Chennai");
		
		wb.write(f);
		f.close();
	}

}
