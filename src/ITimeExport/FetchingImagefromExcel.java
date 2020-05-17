/**
 * 
 */
package ITimeExport;
import java.text.DateFormatSymbols;
import java.util.Calendar;

import javax.imageio.ImageIO;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;

import java.awt.Image;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * @author 10638193
 *
 */
public class FetchingImagefromExcel {

	/**
	 * @param args
	 */
	public static void main(String[] args) throws IOException {
		Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.YEAR, 2020);


        Workbook workbook = new Workbook();
        System.out.println("workbook created");
        //Load the Excel file

         workbook.loadFromFile("C:\\Itime Sheet\\ITimeData.xlsx");
		 Worksheet sheet0 = workbook.getWorksheets().get("Employee");


        // Get the weekday and print it
        int weekday = calendar.get(Calendar.DAY_OF_WEEK);
        System.out.println("Weekday: " + weekday);

        // Get weekday name
        DateFormatSymbols dfs = new DateFormatSymbols();
        System.out.println("Weekday: " + dfs.getWeekdays()[weekday]);
        File imageFile = new File("D:\\java\\NormalImage.png");;
		XMLSlideShow ppt = new XMLSlideShow();  

        if(dfs.getWeekdays()[weekday].toString().equalsIgnoreCase("Tuesday") ||  dfs.getWeekdays()[weekday].toString().equalsIgnoreCase("Thursday")){
              sheet0.saveToImage("D:\\java\\NormalImage.png");
    		  
    		  try (OutputStream os = new FileOutputStream("D:\\PPTX\\Reporting.pptx")) {     
  	            XSLFSlide slide = ppt.createSlide();  
  	            byte[] pictureData = IOUtils.toByteArray(new FileInputStream("D:\\java\\NormalImage.png"));  
  	            XSLFPictureData pd = ppt.addPicture(pictureData, XSLFPictureData.PictureType.PNG);  
  	            XSLFPictureShape pic = slide.createPicture(pd);
  	            
  	            ppt.write(os);  
  	            System.out.println("Image has been written succesfully");
  	        }catch(Exception e) {  
  	            System.out.println(e);  
  	        }  
	          System.out.println("Image converted");
			  System.out.println("Date is correct");
			  System.out.println("Image fetched according to Tuesday or Thursday");
			  System.out.println("Screenshot is taken and written intp PPT");
		}
		else{
	         BufferedImage bufferedImage =null;
			 bufferedImage = ImageIO.read(imageFile);
             BufferedImage image=bufferedImage.getSubimage(300, 150, 200, 200);
             File pathFile = new File("D:\\java\\CroppedImage.png");
             try (OutputStream os = new FileOutputStream("D:\\PPTX\\Reporting.pptx")) {     
 	            XSLFSlide slide = ppt.createSlide();  
 	            byte[] pictureData = IOUtils.toByteArray(new FileInputStream("D:\\java\\CroppedImage.png"));  
 	            XSLFPictureData pd = ppt.addPicture(pictureData, XSLFPictureData.PictureType.PNG);  
 	            XSLFPictureShape pic = slide.createPicture(pd);
 	            
 	            ppt.write(os);  
 	            System.out.println("Image has been written succesfully");
 	        }catch(Exception e) {  
 	            System.out.println(e);  
 	        }  
             ImageIO.write(image,"jpeg", pathFile);
			  System.out.println("Day is not thursday");
			  System.out.println("Image fetched according to other day");
			  System.out.println("Screenshot is taken and written intp PPT in case not Tuesday or Thursday");


		}    

 
	}

}
