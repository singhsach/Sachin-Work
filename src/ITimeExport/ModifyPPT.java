/**
 * 
 */
package ITimeExport;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

/**
 * @author 10638193
 *
 */
public class ModifyPPT {

	/**
	 * @param args
	 * @throws IOException 
	 */
	public static void main(String[] args) throws IOException {
		//creating a presentation
	      XMLSlideShow ppt = new XMLSlideShow();
	      
	      //creating a slide in it
	      XSLFSlide slide = ppt.createSlide();

	      //reading an image
	      File image1 = new File("D:\\java\\image.png");
	      File image2 = new File("D:\\java\\image.png");



	      //converting it into a byte array
	      byte[] picture1 = IOUtils.toByteArray(new FileInputStream(image1));
	      byte[] picture2 = IOUtils.toByteArray(new FileInputStream(image2));



	      //adding the image to the presentation
	      XSLFPictureData pd = ppt.addPicture(picture1, XSLFPictureData.PictureType.PNG);  
          XSLFPictureShape pic = slide.createPicture(pd);
          
          XSLFPictureData pd1 = ppt.addPicture(picture2, XSLFPictureData.PictureType.PNG);  
          XSLFPictureShape pic1 = slide.createPicture(pd1);


	      //creating a file object
	      File file = new File("D:\\PPT\\Reporting.pptx");
	      FileOutputStream out = new FileOutputStream(file);

	      //saving the changes to a file
	      ppt.write(out);
	      
	      System.out.println("image added successfully");
	      out.close();
	   }

}
