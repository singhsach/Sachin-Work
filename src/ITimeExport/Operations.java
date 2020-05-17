package ITimeExport;

import java.io.File;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Operations {
	public void readData(XSSFSheet Sheet1) throws Exception {
		Iterator<Row> rowIterator = Sheet1.iterator();
		ITimeEmp emp = new ITimeEmp();
		int rowcount = 0;
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			if(rowcount!=0) {
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				switch (cell.getCellType()) {
				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						//emp.setDate(cell.getDateCellValue());
					} else 
						if (cell.getColumnIndex() == 3)
							emp.setTime(cell.getNumericCellValue());
						else if(cell.getColumnIndex()==0)
							emp.setPS_id( (long) cell.getNumericCellValue());
					break;
				case STRING:
					 if (cell.getColumnIndex() == 1) {
						emp.setName(cell.getStringCellValue());
					}
					 
					break;
				}
			}
			writeData(emp);
			System.out.println(rowcount);
			}
			rowcount++;
		}
	}

	public void readExcelCappm(XSSFSheet sheet) throws Exception {

		Iterator<Row> rowIterator = sheet.iterator();
		CaPpmEmp emp = new CaPpmEmp();
		int count = 1;
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			if(count!=1) {
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				switch (cell.getCellType()) {
				case BOOLEAN:
					break;
				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell) && cell.getColumnIndex() == 9) {
						emp.setDate(cell.getDateCellValue());
					} else 
						if (cell.getColumnIndex() == 14)
							emp.setTime(cell.getNumericCellValue());
						else if (cell.getColumnIndex() == 4) 
							emp.setPsId((long) cell.getNumericCellValue());
					break;
				case STRING:
					
					 if (cell.getColumnIndex() == 5) {
						emp.setName(cell.getStringCellValue());
					} else if (cell.getColumnIndex() == 7)
						emp.setManagerName(cell.getStringCellValue());
					break;
				}
			}
			writeData(emp, count);
			System.out.println(count);
			}
			count++;
		}
	}

	private void writeData(CaPpmEmp emp, int rowCount) throws Exception {
		File Result = new File("C:\\Itime Sheet\\Result\\Result.xlsx");
		FileInputStream ResultFile = new FileInputStream(Result);
		XSSFWorkbook Resultworkbook = new XSSFWorkbook(ResultFile);
		XSSFSheet ResultSheet = Resultworkbook.getSheetAt(0);
		ResultSheet.getRow(rowCount).createCell(3).setCellValue(emp.getPsId());
		ResultSheet.getRow(rowCount).createCell(4).setCellValue(emp.getName());
		ResultSheet.autoSizeColumn(4);
		
		CellStyle cellStyle = Resultworkbook.createCellStyle();
		CreationHelper createHelper = Resultworkbook.getCreationHelper();
		short dateFormat = createHelper.createDataFormat().getFormat("M-D-YY");
		cellStyle.setDataFormat(dateFormat);
		Cell cell = ResultSheet.getRow(rowCount).createCell(5);
		cell.setCellValue(emp.getDate());
		System.out.println(emp.getDate());
		cell.setCellStyle(cellStyle);
		ResultSheet.getRow(rowCount).createCell(6).setCellValue(emp.getTime());
		ResultSheet.getRow(rowCount).createCell(7).setCellValue(emp.getManagerName());
		ResultSheet.autoSizeColumn(7);

		try {
			FileOutputStream out = new FileOutputStream(Result);
			Resultworkbook.write(out);
			out.close();
			System.out.println("Excel written successfully..");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public void writeData(ITimeEmp emp) throws Exception {
		
		File Result = new File("C:\\Itime Sheet\\Result\\Result.xlsx");
		FileInputStream ResultFile = new FileInputStream(Result);
		XSSFWorkbook Resultworkbook = new XSSFWorkbook(ResultFile);
		XSSFSheet ResultSheet = Resultworkbook.getSheetAt(0);
		Row DataRow;
		

		int rowCount = ResultSheet.getLastRowNum();
		DataRow = ResultSheet.createRow(rowCount + 1);

		DataRow.createCell(0).setCellValue(emp.getPS_id());
		DataRow.createCell(1).setCellValue(emp.getName());
		ResultSheet.autoSizeColumn(1);
//		CellStyle cellStyle = Resultworkbook.createCellStyle();
//		CreationHelper createHelper = Resultworkbook.getCreationHelper();
//		short dateFormat = createHelper.createDataFormat().getFormat("M-D-YY");
//		cellStyle.setDataFormat(dateFormat);
//		Cell cell = DataRow.createCell(2);
//		cell.setCellValue(emp.getDate());
//		System.out.println(emp.getDate());
//		cell.setCellStyle(cellStyle);

		DataRow.createCell(2).setCellValue(emp.getTime());
		try {
			FileOutputStream out = new FileOutputStream(Result);
			Resultworkbook.write(out);
			out.close();
			System.out.println("Excel written successfully..");

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public int CompareExcel() throws Exception {

		File Result = new File("C:\\Itime Sheet\\Result\\Result.xlsx");
		FileInputStream ResultFile = new FileInputStream(Result);
		XSSFWorkbook Resultworkbook = new XSSFWorkbook(ResultFile);
		XSSFSheet ResultSheet = Resultworkbook.getSheetAt(0);
		Operations operations = new Operations();
		int errorCount=0;
		Iterator<Row> rowIterator = ResultSheet.iterator();
		ITimeEmp emp = new ITimeEmp();
		CaPpmEmp caPpmEmp = new CaPpmEmp();
		int rowcount = 0;

		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();

			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				switch (cell.getCellType()) {
				case BOOLEAN:
					break;
				case NUMERIC:
//					if (DateUtil.isCellDateFormatted(cell) && cell.getColumnIndex() == 2) {
//						emp.setDate(cell.getDateCellValue());

				//	}else
//					if (DateUtil.isCellDateFormatted(cell) && cell.getColumnIndex() == 6) {
//						caPpmEmp.setDate(cell.getDateCellValue());
//					} else {
						if (cell.getColumnIndex() == 2) {
							emp.setTime(cell.getNumericCellValue());
						} else if (cell.getColumnIndex() == 6) {
							caPpmEmp.setTime(cell.getNumericCellValue());
						}
						else if (cell.getColumnIndex() == 3) {
							caPpmEmp.setPsId((long) cell.getNumericCellValue());
						}
						if (cell.getColumnIndex() == 0) 
							emp.setPS_id((long) cell.getNumericCellValue());
					//}
					break;
				case STRING:
				
					 if (cell.getColumnIndex() == 1) {
						emp.setName(cell.getStringCellValue());
					} else if (cell.getColumnIndex() == 4) {
						caPpmEmp.setName(cell.getStringCellValue());

					}
					break;
				}

			}
			if (rowcount == 0 || rowcount==1) {
			} else
				errorCount=errorCount+operations.match(emp, caPpmEmp, rowcount, ResultSheet, Result, Resultworkbook);
			rowcount++;
		}
		
		return errorCount;
	}

	public int match(ITimeEmp emp, CaPpmEmp caPpmEmp, int rowcount, XSSFSheet ResultSheet, File Result,
			XSSFWorkbook Resultworkbook) throws Exception {
		CellStyle style1 = Resultworkbook.createCellStyle();
		style1.setFillForegroundColor(IndexedColors.RED.getIndex());
		style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style1.setBorderBottom(BorderStyle.THIN);
		style1.setBorderLeft(BorderStyle.THIN);
		style1.setBorderTop(BorderStyle.THIN);
		style1.setBorderRight(BorderStyle.THIN);

		CellStyle style2 = Resultworkbook.createCellStyle();
		style2.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style2.setBorderBottom(BorderStyle.THIN);
		style2.setBorderLeft(BorderStyle.THIN);
		style2.setBorderTop(BorderStyle.THIN);
		style2.setBorderRight(BorderStyle.THIN);
		ResultSheet.autoSizeColumn(9);
		FileOutputStream fop;
		int flag=0;
		if (emp.getName().equals(caPpmEmp.getName())) {
			if (emp.getPS_id()==caPpmEmp.getPsId()) {
				//if (emp.getDate().compareTo(caPpmEmp.getDate()) == 0) {
					if (emp.getTime() == caPpmEmp.getTime()) {
						ResultSheet.getRow(rowcount).createCell(9).setCellValue("match");
						ResultSheet.getRow(rowcount).getCell(9).setCellStyle(style2);
						System.out.println("match");

					} else {
						ResultSheet.getRow(rowcount).createCell(9).setCellValue("Time mis-match");
						ResultSheet.getRow(rowcount).getCell(9).setCellStyle(style1);
						ResultSheet.getRow(rowcount).getCell(2).setCellStyle(style1);
						ResultSheet.getRow(rowcount).getCell(6).setCellStyle(style1);
						flag=1;
						System.out.println("Time mismatch");

					}

//				} else {
//					ResultSheet.getRow(rowcount).createCell(9).setCellValue("Date mis-match");
//					ResultSheet.getRow(rowcount).getCell(9).setCellStyle(style1);
//					ResultSheet.getRow(rowcount).getCell(2).setCellStyle(style1);
//					ResultSheet.getRow(rowcount).getCell(6).setCellStyle(style1);
//					flag=1;
//					System.out.println("date mismatch");
//
//				}

			} else {
				ResultSheet.getRow(rowcount).createCell(9).setCellValue("Ps id mis-match");
				ResultSheet.getRow(rowcount).getCell(9).setCellStyle(style1);
				ResultSheet.getRow(rowcount).getCell(0).setCellStyle(style1);
				ResultSheet.getRow(rowcount).getCell(3).setCellStyle(style1);
				flag=1;
				System.out.println("ps id  mismatch");

			}

		} else {
			ResultSheet.getRow(rowcount).createCell(9).setCellValue("Name mis-match");
			ResultSheet.getRow(rowcount).getCell(9).setCellStyle(style1);
			ResultSheet.getRow(rowcount).getCell(1).setCellStyle(style1);
			ResultSheet.getRow(rowcount).getCell(4).setCellStyle(style1);
			flag=1;
			System.out.println("name mismatch");

		}
		
		fop = new FileOutputStream(Result);
		Resultworkbook.write(fop);
		rowcount++;
		return flag;

	}
	public void CreateHeaderCol() throws Exception
	{
		File Result = new File("C:\\Itime Sheet\\Result\\Result.xlsx");
		FileInputStream ResultFile = new FileInputStream(Result);
		XSSFWorkbook Resultworkbook = new XSSFWorkbook(ResultFile);
		XSSFSheet ResultSheet = Resultworkbook.getSheetAt(0);
		CellStyle style1 = Resultworkbook.createCellStyle();
		style1.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style1.setBorderBottom(BorderStyle.THIN);
		style1.setBorderLeft(BorderStyle.THIN);
		style1.setBorderTop(BorderStyle.THIN);
		style1.setBorderRight(BorderStyle.THIN);
		style1.setAlignment(HorizontalAlignment.CENTER);
		ResultSheet.createRow(0).createCell(0).setCellValue("I Time Data");
		ResultSheet.getRow(0).createCell(4).setCellValue("CA-PPM Data");
		ResultSheet.getRow(0).getCell(0).setCellStyle(style1);
		ResultSheet.getRow(0).getCell(4).setCellStyle(style1);
		ResultSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));
		ResultSheet.addMergedRegion(new CellRangeAddress(0, 0, 3, 7));
		//ResultSheet.getRow(0).getCell(0).set
		ResultSheet.createRow(1).createCell(0).setCellValue("PS Number");
		//ResultSheet.getRow(1).setRowStyle(style1);
		ResultSheet.getRow(1).createCell(1).setCellValue("Name");
		//ResultSheet.getRow(1).createCell(2).setCellValue("Time Booking Date");
		ResultSheet.getRow(1).createCell(2).setCellValue("Billable hours");
		ResultSheet.getRow(1).createCell(3).setCellValue("Resource Id");
		ResultSheet.getRow(1).createCell(4).setCellValue("Resource Name");
		ResultSheet.getRow(1).createCell(5).setCellValue("Date Worked");
		ResultSheet.getRow(1).createCell(6).setCellValue("Resource Id");
		ResultSheet.getRow(1).createCell(7).setCellValue("Resource Manager");
		ResultSheet.getRow(1).createCell(9).setCellValue("Result");
		
		try {
			FileOutputStream out = new FileOutputStream(Result);
			Resultworkbook.write(out);
			out.close();
			System.out.println("Excel written successfully..");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
