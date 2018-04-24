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

public class WriteExcel1 {
	
	public static void main(String[] args) {
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Data 2");
		
		Map<String, Object[]> data = new HashMap<String, Object[]>();
		data.put("1", new Object[] {"numberSetOne", "numberSetTwo", "wordSetOne"});
		data.put("2", new Object[] {60d, 10d, "With"});
		data.put("3", new Object[] {38d, 5d, "Best"});
		data.put("4", new Object[] {2095d, 8d, "Like"});
		data.put("5", new Object[] {145d, 512d, "Rest"});
		
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
			FileOutputStream out = new FileOutputStream(new File("C:\\data2.xls"));
			
			workbook.write(out);
			
			out.close();
			System.out.println("Second Excel sheet created!");
		}
		catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} 
	}
	
}
