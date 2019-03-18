package com.example.filedemo.operation;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.example.filedemo.model.Transcationfield;

public class Writefile {

	private String[] columns = { "EmpNumber", "Add1", "Add2", "Street", "city", "mobile1", "mobile2", "landline", "DOB",
			"DepCode", "Designation", "DepartmentCode", "Location" };
	
	private String[] columnsError = { "EmpNumber", "Add1", "Add2", "Street", "city", "mobile1", "mobile2", "landline", "DOB",
			"DepCode", "Designation", "DepartmentCode", "Location" };

	private List<Transcationfield> employees = new ArrayList<>();

	public List<Transcationfield> newFilegeneration(List<Transcationfield> TransList,int startSize,int skipSize, String sheetName) throws IOException {

		Workbook workbook = new XSSFWorkbook();
		
		CreationHelper createHelper = workbook.getCreationHelper();

		// Create a Sheet
		Sheet sheet = workbook.createSheet("Employee");

		System.out.println("******fine******" + TransList.size());
		// Create a Font for styling header cells
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.RED.getIndex());

		// Create a CellStyle with the font
		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);

		// Create a Row
		Row headerRow = sheet.createRow(0);

		// Creating cells
		for (int i = 0; i < columns.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(columns[i]);
			cell.setCellStyle(headerCellStyle);
		}

		// Cell Style for formatting Date
		CellStyle dateCellStyle = workbook.createCellStyle();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

		// Create Other rows and cells with employees data
		int rowNum = 1;
		
			System.out.println("TransList.size()=="+TransList.size());
					 System.out.println("startsize==="+startSize);
					 System.out.println("endsize==="+skipSize);

			for (Transcationfield ts:TransList.subList(startSize, skipSize))
		    {
			Row row = sheet.createRow(rowNum++);

			row.createCell(0).setCellValue(ts.getEmpNumber());
		
			row.createCell(1).setCellValue(ts.getEmpTransadd1());

			row.createCell(2).setCellValue(ts.getEmpTransadd2());

			row.createCell(3).setCellValue(ts.getEmpTransstreet());

			row.createCell(4).setCellValue(ts.getEmpTranscity());
			
			row.createCell(5).setCellValue( ts.getEmpContactTranscationmobile1());

			row.createCell(6).setCellValue(ts.getEmpContactTranscationmobile2());

			row.createCell(7).setCellValue( ts.getEmpContactTranscationlandline());
			
			row.createCell(8).setCellValue( ts.getEmpDateofbirth());

			row.createCell(9).setCellValue( ts.getEmpDepartmentCode());
			
			row.createCell(10).setCellValue( ts.getEmpDistignation());
			
			row.createCell(11).setCellValue( ts.getEmpDepartmentName());
			
			row.createCell(12).setCellValue( ts.getEmpLocation());
			
			//row.createCell(14).setCellValue( ts.getEmpTransactionerror());

		
	}

		for (int i = 0; i < columns.length; i++) {
			sheet.autoSizeColumn(i);
		}

		// Write the output to a file
		FileOutputStream fileOut;

		fileOut = new FileOutputStream("C:\\Users\\10646744\\Desktop\\AmazonSharefolder\\"+sheetName);
		workbook.write(fileOut);
		fileOut.close();
		workbook.close();

		return employees;

	}
	
	public List<Transcationfield> newFilegeneration(List<Transcationfield> TransList) throws IOException {

		Workbook workbook = new XSSFWorkbook();
		
		CreationHelper createHelper = workbook.getCreationHelper();

		// Create a Sheet
		Sheet sheet = workbook.createSheet("Employee");

		System.out.println("******fine******" + TransList.size());
		// Create a Font for styling header cells
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.RED.getIndex());

		// Create a CellStyle with the font
		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);

		// Create a Row
		Row headerRow = sheet.createRow(0);

		// Creating cells
		for (int i = 0; i < columnsError.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(columnsError[i]);
			cell.setCellStyle(headerCellStyle);
		}

		// Cell Style for formatting Date
		CellStyle dateCellStyle = workbook.createCellStyle();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

		// Create Other rows and cells with employees data
		int rowNum = 1;
		
			System.out.println("TransList.size()=="+TransList.size());
					// System.out.println("startsize==="+startSize);
					// System.out.println("endsize==="+skipSize);

			for (Transcationfield ts:TransList)
		    {
			Row row = sheet.createRow(rowNum++);

			row.createCell(0).setCellValue(ts.getEmpNumber());
		
			row.createCell(1).setCellValue(ts.getEmpTransadd1());

			row.createCell(2).setCellValue(ts.getEmpTransadd2());

			row.createCell(3).setCellValue(ts.getEmpTransstreet());

			row.createCell(4).setCellValue(ts.getEmpTranscity());
			
			row.createCell(5).setCellValue( ts.getEmpContactTranscationmobile1());

			row.createCell(6).setCellValue(ts.getEmpContactTranscationmobile2());

			row.createCell(7).setCellValue( ts.getEmpContactTranscationlandline());
			
			row.createCell(8).setCellValue( ts.getEmpDateofbirth());

			row.createCell(9).setCellValue( ts.getEmpDepartmentCode());
			
			row.createCell(10).setCellValue( ts.getEmpDistignation());
			
			row.createCell(11).setCellValue( ts.getEmpDepartmentName());
			
			row.createCell(12).setCellValue( ts.getEmpLocation());
			
			row.createCell(14).setCellValue( ts.getEmpTransactionerror());

		
	}

		for (int i = 0; i < columnsError.length; i++) {
			sheet.autoSizeColumn(i);
		}

		// Write the output to a file
		FileOutputStream fileOut;

		fileOut = new FileOutputStream("C:\\Users\\10646744\\Desktop\\AmazonSharefolder1\\worksheet.csv");
		workbook.write(fileOut);
		fileOut.close();
		workbook.close();

		return employees;

	}


}
