/**
 * 
 */
package ITimeExport;
import com.aspose.cells.*;
import java.io.*;

/**
 * @author 10638193
 *
 */
public class ConverttoPDF {

	/**
	 * @param args
	 * @throws Exception 
	 */
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub

		try{
			
		Workbook workbook=new Workbook("C:\\Itime Sheet\\ITimeData.xlsx");
		workbook.save("C:\\Excel toPDF\\Excel2PDF_Output.pdf", FileFormatType.PDF);
		System.out.println("Done Successfully");
		
	}
		
		catch(IOException e){
			e.printStackTrace();
		}

}
	
}
