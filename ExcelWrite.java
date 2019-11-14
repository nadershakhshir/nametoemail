package com.nadir1.java1;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {
	
	  public static void main(String[] args) throws IOException, 
		InvalidFormatException {
		  String name = "";
			try{
				FileInputStream file = new FileInputStream(new File("D:\\contacts1.xlsx"));
				
			    XSSFWorkbook workbook1 = new XSSFWorkbook(file);
			 
			    XSSFSheet sheet1 = workbook1.getSheetAt(0);
			     
			    Iterator<Row> rowIterator = sheet1.iterator();
			    while(rowIterator.hasNext()) {
			        Row row = rowIterator.next();
		             ArrayList<String> list = new ArrayList<String>();

			        Iterator<Cell> cellIterator = row.cellIterator();
			        
			        while(cellIterator.hasNext()) {
			            Cell cell = cellIterator.next();
			            switch(cell.getCellType()) {
			             
			                case Cell.CELL_TYPE_STRING:
			                	name = cell.getStringCellValue();
			                	list.add(name);
			                	cell.setCellValue(name);
			                	String joined2 = String.join("",list);
			                    System.out.print(joined2);
			                	cell.setCellValue(joined2+"@debugnomad.com");
			                    break;
			            }
			        }
			        System.out.println("");
			    }
			    file.close();
			    FileOutputStream out =
			    new FileOutputStream(new File("D:\\test.xlsx"));
			    workbook1.write(out);
			    out.close();
			} catch (FileNotFoundException e) {
			    e.printStackTrace();
			} catch (IOException e) {
			    e.printStackTrace();
			}
	}
	  
	}
