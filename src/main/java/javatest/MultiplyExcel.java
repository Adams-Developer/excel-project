package javatest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;



public class MultiplyExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		//get the mergeExcel sheet
		String excelFile = "C:\\mergedSheets.xls";
		
		FileInputStream  inputStream = new FileInputStream(new File(excelFile));
		
		//create workbook
		HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
		Sheet sheet = workbook.getSheetAt(0);
		
		//get the sum of rows multiplied					
		int rowCount = 0; 
		int columnCount = 0;
		
		Row rowTotal = sheet.createRow(columnCount +  12);
		Cell cellName = rowTotal.createCell(0);
		
		cellName.setCellValue("Multiplication Total:");
		
		Cell cellTotal = rowTotal.createCell(1);
		Cell cellTotal1 = rowTotal.createCell(2);
		Cell cellTotal2 = rowTotal.createCell(3);
		Cell cellTotal3 = rowTotal.createCell(4);

		//cellTotal.setCellFormula("SUMPRODUCT(A2:A5,A7:A10)");
		cellTotal.setCellFormula("SUMPRODUCT(A2,A7)");
		cellTotal1.setCellFormula("SUMPRODUCT(A3,A8)");
		cellTotal2.setCellFormula("SUMPRODUCT(A4,A9)");
		cellTotal3.setCellFormula("SUMPRODUCT(A5,A10)");	

		//divide rows from data sets		
		Row rowTotal1 = sheet.createRow(rowCount + 13);
		Cell cellName1 = rowTotal1.createCell(0);
		cellName1.setCellValue("Division Total:");
		
		Cell cellTotal4 = rowTotal1.createCell(1);
		Cell cellTotal5 = rowTotal1.createCell(2);
		Cell cellTotal6 = rowTotal1.createCell(3);
		Cell cellTotal7 = rowTotal1.createCell(4);
		
		cellTotal4.setCellFormula("(B2/B7)");
		cellTotal5.setCellFormula("(B3/B8)");
		cellTotal6.setCellFormula("(B4/B9)");
		cellTotal7.setCellFormula("(B5/B10)");
		
		//concatenate words
		Row rowTotal2 = sheet.createRow(rowCount + 14);
		Cell cellName2 = rowTotal2.createCell(0);
		cellName2.setCellValue("Words:");
		
		Cell cellTotal8 = rowTotal2.createCell(1);
		Cell cellTotal9 = rowTotal2.createCell(2);
		Cell cellTotal10 = rowTotal2.createCell(3);
		Cell cellTotal11 = rowTotal2.createCell(4);
		
		cellTotal8.setCellFormula("(C2 & C7)");
		cellTotal9.setCellFormula("(C3 & C8 )" );
		cellTotal10.setCellFormula("(C4 & C9 )" );
		cellTotal11.setCellFormula("(C5 & C10 )" );
		
		ArrayList<Integer> totals = new ArrayList<>();
		totals.add(2400);
		totals.add(7600);
		totals.add(150840);
		totals.add(14355);

		System.out.println(totals);
		
		ArrayList<Integer> totals1 = new ArrayList<>();
		totals1.add(8);
		totals1.add(40);
		totals1.add(2);
		totals1.add(2);

		System.out.println(totals1);
		
		ArrayList<String> totals2 = new ArrayList<>();
		totals2.add("Mess With");
		totals2.add("The Best");
		totals2.add("Die Like");
		totals2.add("The Rest");

		System.out.println(totals2);
		
		inputStream.close();
		
		FileOutputStream outputStream = new FileOutputStream(excelFile);
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();
	}

}
