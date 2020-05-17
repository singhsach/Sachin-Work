package ITimeExport;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	public int readExcel() throws Exception {

		File iTime = new File("C:\\Itime Sheet\\ITimeData.xlsx");
		FileInputStream fis = new FileInputStream(iTime);
		XSSFWorkbook workbook1 = new XSSFWorkbook(fis);
		XSSFSheet Sheet1 = workbook1.getSheetAt(0);

		File CaPpm = new File("C:\\Itime Sheet\\ca-ppm.xlsx");
		FileInputStream fis1 = new FileInputStream(CaPpm);
		XSSFWorkbook workbook2 = new XSSFWorkbook(fis1);
		XSSFSheet Sheet2 = workbook2.getSheetAt(0);
		
//		
//		File Result = new File("C:\\Itime Sheet\\Result\\Result.xlsx");
//		FileInputStream ResultFile = new FileInputStream(Result);
		XSSFWorkbook Resultworkbook = new XSSFWorkbook();
		OutputStream os=new FileOutputStream("C:\\Itime Sheet\\Result\\Result.xlsx");
		XSSFSheet SheetResult = Resultworkbook.createSheet("Result1");
		Resultworkbook.write(os);

//		XSSFSheet ResultSheet = Resultworkbook.createSheet("C:\\Itime Sheet\\Result\\Result.xlsx");
		
		Operations operations = new Operations();
		operations.CreateHeaderCol();
		operations.readData(Sheet1);
		operations.readExcelCappm(Sheet2);
		int errorcount =operations.CompareExcel();
		System.out.println(errorcount);
		return errorcount;
		
		//workbook1.close();
		//workbook2.close();
		//Resultworkbook.close();

	}
}
