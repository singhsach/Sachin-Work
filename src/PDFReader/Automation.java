/**
 * 
 */
package PDFReader;



import java.awt.Robot;
import java.io.File;
import java.io.FileInputStream;
import java.io.Reader;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import javax.imageio.ImageIO;

import com.java.Email.MailReport;

import FetchingImageFromExcel.FetchingImagefromExcel;
import createPPT.Activity;

/**
 * @author 10638193
 *
 */
public class Automation {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		try{
			FetchingImagefromExcel p1=new FetchingImagefromExcel();
			PDF2PPT p2=new PDF2PPT();
             Activity act = new Activity();
            MailReport mail=new MailReport();

            p1.ImageFromExcel();
			p2.convertPDF2Image("D:/PDF");
			act.logIn();
			
			p2.convertPDF2PPT("D:/Image");
			
			mail.sendMail();
			
			
			

			
			
			
			
	//		p1.convertPDF2PPT("D:/PDFimages");
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}

}
