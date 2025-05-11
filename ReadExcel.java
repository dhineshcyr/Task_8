package Task_8;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

		public static void main(String[] args) {

			ReadExcel rd=new ReadExcel();
			
			for(int i=0;i<4;i++) {
				
				for(int j=0;j<3;j++ ) {
				
					System.out.print(rd.getReadData("sheet1", i, j)+"  ");
				}
				System.out.println();
			}
			
			
		  rd.writeexcel("sheet1", 4, 0, "abi");
		  rd.writeexcel("sheet1", 4, 1, "23");
		  rd.writeexcel("sheet1", 4, 2, "abi@gmail.com");
		  
			
		}
	
		public void writeexcel(String sheetName,int rowNum,int cellNum,String desc) {
			FileInputStream fis;
			XSSFWorkbook wb;
		
			try {
				
				fis=new FileInputStream("C:\\Users\\mdhin\\eclipse-workspace\\GuviTask\\src\\Task_8\\stud.xlsx");
				wb=new XSSFWorkbook(fis);
				XSSFSheet s=wb.getSheet(sheetName);
				XSSFRow r=s.getRow(rowNum);
				XSSFCell c=r.createCell(cellNum);
			c.setCellValue(desc);
			
			FileOutputStream fos=new FileOutputStream("C:\\Users\\mdhin\\eclipse-workspace\\GuviTask\\src\\Task_8\\stud.xlsx");
			wb.write(fos);
			
			} catch (FileNotFoundException e) {
				
				e.printStackTrace();
			} catch(IOException e) {
				e.printStackTrace();
			}
		
			
			
		}
		public String getReadData(String sheetName,int rowNum,int colNum) {
			
			String val=null;
	     try {
	    	
			FileInputStream fis=new  FileInputStream("C:\\Users\\mdhin\\eclipse-workspace\\GuviTask\\src\\Task_8\\stud.xlsx");
			XSSFWorkbook wb=new XSSFWorkbook(fis); 
			XSSFSheet s=wb.getSheet(sheetName); 
			XSSFRow r=s.getRow(rowNum); 
			XSSFCell c=r.getCell(colNum);  
		    val=ReadExcel.getCellValue(c); 
			
			fis.close();
		
			wb.close();
			
		} catch (FileNotFoundException e) {
		
			e.printStackTrace();
		}catch  (IOException e) {
			e.printStackTrace();
		}
	     return val;
		}
	
	  public static  String getCellValue(XSSFCell c)
	  {
		  switch(c.getCellType()) {   
		  case NUMERIC:
			  return String.valueOf(c.getNumericCellValue()); 
		  case BOOLEAN:
			  return String.valueOf(c.getBooleanCellValue());
		  case STRING:
			  return c.getStringCellValue(); 
			  default:
				  return c.getStringCellValue();
		  
		  }
	  }

}


//OutPut
//NAME  AGE  EMAIL  
//guna  24.0  guna@gmail.com  
//sri  45.0  sri@gmail.com  
//jash  34.0  jash@gmail.com  