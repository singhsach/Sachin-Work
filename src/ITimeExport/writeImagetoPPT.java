/**
 * 
 */
package ITimeExport;



 

import java.io.FileInputStream;  
import java.io.FileNotFoundException;  
import java.io.FileOutputStream;  
import java.io.IOException;  
import java.io.OutputStream;  
import org.apache.poi.util.IOUtils;  
import org.apache.poi.xslf.usermodel.XMLSlideShow;  
import org.apache.poi.xslf.usermodel.XSLFPictureData;  
import org.apache.poi.xslf.usermodel.XSLFPictureShape;  
import org.apache.poi.xslf.usermodel.XSLFSlide;

/**
 * @author 10638193
 *
 */
public class writeImagetoPPT {

	/**
	 * @param args
	 */
	public static void main(String args[]) {
		   
		  XMLSlideShow ppt = new XMLSlideShow();  
	        try (OutputStream os = new FileOutputStream("D:\\PPT\\Reporting.pptx")) {     
	            XSLFSlide slide = ppt.createSlide();  
	            byte[] pictureData = IOUtils.toByteArray(new FileInputStream("D:\\java\\image.png"));  
	            XSLFPictureData pd = ppt.addPicture(pictureData, XSLFPictureData.PictureType.PNG);  
	            XSLFPictureShape pic = slide.createPicture(pd);
	            
	            ppt.write(os);  
	            System.out.println("Image has been written succesfully");
	        }catch(Exception e) {  
	            System.out.println(e);  
	        }  
	    }  

}
