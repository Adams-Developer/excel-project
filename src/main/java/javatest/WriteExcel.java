package javatest;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

public class WriteExcel {
	
	public static void main(String[] args) {
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Data 1");
		
		Map<String, Object[]> data = new HashMap<String, Object[]>();
		data.put("1", new Object[] {"numberSetOne", "numberSetTwo", "wordSetOne"});
		data.put("2", new Object[] {40d, 80d, "Mess"});
		data.put("3", new Object[] {200d, 200d, "The"});
		data.put("4", new Object[] {72d, 16d, "Die"});
		data.put("5", new Object[] {99d, 1024d, "The"});
		
		Set<String> keyset = data.keySet();
		
		int rowNum = 0;
		
		for (String key : keyset) {
			
			Row row = sheet.createRow(rowNum++);
			
			Object [] objArray = data.get(key);
			
			int cellNum = 0;
			
			for (Object obj : objArray) {
				
				Cell cell = row.createCell(cellNum++);
				
				if (obj instanceof String)
					cell.setCellValue((String)obj);
				else if (obj instanceof Double)
					cell.setCellValue((Double)obj);
			}
		}
		
		try {
			FileOutputStream out = new FileOutputStream(new File("C:\\data1.xls"));
			
			workbook.write(out);
			
			out.close();
			System.out.println("Excel sheet created!");
		}
		catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} 
	}
	
}

