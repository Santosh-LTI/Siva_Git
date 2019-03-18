package com.example.filedemo.operation;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Autowired;

import com.example.filedemo.model.TransTable;
import com.example.filedemo.model.Transcationfield;

//@ConfigurationProperties("bean")
//@ConfigurationProperties(prefix = "bean")
public class Readfile {

	@Autowired
	private Transcationfield tf;
	String cellValue;
	String datePattern = "\\d{1,2}/\\d{1,2}/\\d{4}";
	Writefile wf = new Writefile();
	int startSize;
	int startvalue = 2;
	int skipSize = 4;
	int endSize;
	boolean flag = false;
	TransTable ttable;
	//private String path;

	List<TransTable> tempList = new ArrayList<>();

	
	public List excelRead(String filePath, String fileName)
			throws EncryptedDocumentException, InvalidFormatException, IOException {
		List<Transcationfield> transList = new ArrayList<Transcationfield>();
		// File file1 = new
		// File("C:\\Users\\10646744\\Desktop\\AmazonSharefolder1");
		// if (!file1.exists()) {
		// file1.mkdir();
		// }
		// Creating a Workbook from an Excel file (.xls or .xlsx)
		Workbook workbook = WorkbookFactory.create(new File(filePath + "\\" + fileName));

		// Retrieving the number of sheets in the Workbook
		System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

		workbook.forEach(sheet -> {
			System.out.println("=> " + sheet.getSheetName());
		});

		// Getting the Sheet at index zero
		Sheet sheet = workbook.getSheetAt(0);

		// Create a DataFormatter to format and get each cell's value as String
		DataFormatter dataFormatter = new DataFormatter();

		// you can use Java 8 forEach loop with lambda

		System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
		sheet.forEach(row -> {
			tf = new Transcationfield();
			row.forEach(cell -> {
				String[] s = new String[10];
				if (cell.getColumnIndex() == 0) {
					// ar[row]=printCellValue(cell);
					tf.setEmpNumber(printCellValue(cell));
					System.out.println("cellvalue==" + (printCellValue(cell)));
				}
				if (cell.getColumnIndex() == 1) {
					tf.setEmpTransadd1(printCellValue(cell));

				}
				if (cell.getColumnIndex() == 2) {
					tf.setEmpTransadd2(printCellValue(cell));

				}
				if (cell.getColumnIndex() == 3) {
					tf.setEmpTransstreet(printCellValue(cell));

				}
				if (cell.getColumnIndex() == 4) {
					tf.setEmpTranscity(printCellValue(cell));

				}
				if (cell.getColumnIndex() == 5) {
					tf.setEmpContactTranscationmobile1(printCellValue(cell));

				}
				if (cell.getColumnIndex() == 6) {
					tf.setEmpContactTranscationmobile2(printCellValue(cell));

				}
				if (cell.getColumnIndex() == 7) {
					tf.setEmpContactTranscationlandline(printCellValue(cell));

				}
				if (cell.getColumnIndex() == 8) {
					tf.setEmpDateofbirth(printCellValue(cell));

				}
				if (cell.getColumnIndex() == 9) {
					tf.setEmpDepartmentCode(printCellValue(cell));

				}
				if (cell.getColumnIndex() == 10) {
					tf.setEmpDistignation(printCellValue(cell));

				}
				if (cell.getColumnIndex() == 11) {
					tf.setEmpDepartmentName(printCellValue(cell));

				}
				if (cell.getColumnIndex() == 12) {
					tf.setEmpLocation(printCellValue(cell));

				}
				if (cell.getColumnIndex() == 13) {
					tf.setEmpLocation(printCellValue(cell));

				}
				if (cell.getColumnIndex() == 14) {
					tf.setEmpLocation(printCellValue(cell));

				}if (cell.getColumnIndex() == 15) {
					tf.setEmpLocation(printCellValue(cell));

				}if (cell.getColumnIndex() == 16) {
					tf.setEmpLocation(printCellValue(cell));

				}if (cell.getColumnIndex() == 17) {
					tf.setEmpLocation(printCellValue(cell));

				}if (cell.getColumnIndex() == 182) {
					tf.setEmpLocation(printCellValue(cell));

				}if (cell.getColumnIndex() == 19) {
					tf.setEmpLocation(printCellValue(cell));

				}if (cell.getColumnIndex() == 20) {
					tf.setEmpLocation(printCellValue(cell));

				}if (cell.getColumnIndex() == 21) {
					tf.setEmpLocation(printCellValue(cell));

				}if (cell.getColumnIndex() == 22) {
					tf.setEmpLocation(printCellValue(cell));

				}if (cell.getColumnIndex() == 23) {
					tf.setEmpLocation(printCellValue(cell));

				}if (cell.getColumnIndex() == 24) {
					tf.setEmpLocation(printCellValue(cell));

				}if (cell.getColumnIndex() == 25) {
					tf.setEmpLocation(printCellValue(cell));

				}if (cell.getColumnIndex() == 26) {
					tf.setEmpLocation(printCellValue(cell));

				}if (cell.getColumnIndex() == 27) {
					tf.setEmpLocation(printCellValue(cell));

				}if (cell.getColumnIndex() == 28) {
					tf.setEmpLocation(printCellValue(cell));

				}if (cell.getColumnIndex() == 29) {
					tf.setEmpLocation(printCellValue(cell));

				}

			});
			transList.add(tf);

			// System.out.println();

		});
		// System.out.println("size of array"+transList.size());

//		for()
//		if(trranscation=="")
//		{}
		
		for (Transcationfield s : transList.subList(2, transList.size())) {
			String dob = (s.getEmpDateofbirth() == null ? "" : s.getEmpDateofbirth());
			boolean isDate = dob.matches(datePattern) || (dob.matches(""));
			System.out.println("idate==" + isDate);
			System.out.println("idob==" + dob);

			if (isDate == true) {
				System.out.println("*****ok*****");
				s.setEmpTransactionerror("");
				System.out.println("s.getTranscationError()==" + s.getEmpTransactionerror());

			} else {
				System.out.println("*****NOTok*****");
				s.setEmpTransactionerror("DOB  Wrong");
				flag = true;

			}
			System.out.println("value==" + s.toString());

		}

		System.out.println("***************");

		// Closing the workbook
		workbook.close();
		System.out.println("flag===" + flag);
		// if (flag==true) {
		// wf.newFilegeneration(transList);
		// System.out.println("flag is true");
		// } else {
		System.out.println("flag is false");
		endSize = (transList.size() / 2);
		for (startSize = 1; startSize < endSize; startSize++) {
			String sheetName = "worksheet" + (startSize - 1) + ".csv";
			System.out.println("startvalue==" + startvalue);
			System.out.println("skip==" + skipSize);
			wf.newFilegeneration(transList, startvalue, skipSize, sheetName);

			ttable = new TransTable();
			ttable.setStatus("O");
			ttable.setTeam("RPA");
			// ttable.setDate(String.valueOf(LocalDateTime.now()));
			ttable.setFileName(sheetName);
			System.out.println("transList2===" + tempList.size());
			tempList.add(ttable);

			startvalue = startvalue + 2;
			System.out.println("startSize11==" + startvalue);
			// if(startSize+2>=endSize){
			skipSize = skipSize + 2;
			System.out.println("startvalue2==" + startvalue);
			System.out.println("skip2==" + skipSize);
			// }

		}
		return tempList;

	}

	private String printCellValue(Cell cell) {
		switch (cell.getCellTypeEnum()) {
		case BOOLEAN:
			// System.out.print(cell.getBooleanCellValue());
			cellValue = String.valueOf(cell.getBooleanCellValue());
			break;
		case STRING:
			// System.out.print(cell.getRichStringCellValue().getString());
			cellValue = cell.getRichStringCellValue().getString();

			break;
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				// System.out.print(cell.getDateCellValue());
				cellValue = String.valueOf(cell.getDateCellValue());

			} else {
				// System.out.print(cell.getNumericCellValue());
				cellValue = String.valueOf(cell.getNumericCellValue());

			}
			break;
		case FORMULA:
			// System.out.print(cell.getCellFormula());
			cellValue = String.valueOf(cell.getCellFormula());

			break;
		case BLANK:
			// System.out.print("----");
			cellValue = String.valueOf("");

			break;
		default:
			System.out.print("");
		}

		System.out.print("\t");
		return cellValue;
	}

}
