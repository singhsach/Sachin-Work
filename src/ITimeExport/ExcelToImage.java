package ITimeExport;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

import javax.imageio.ImageIO;

import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;

public class ExcelToImage {
	  public static void main(String[] args) throws IOException {
		  
		          //Create a workbook instance
		  
		          Workbook workbook = new Workbook();
		          System.out.println("workbook created");
		          //Load the Excel file
		  
		          workbook.loadFromFile("C:\\Itime Sheet\\ITimeData.xlsx");
		  System.out.println("sheet opened");
		   
		  
		          //Get the first worksheet
		  
		  Worksheet sheet0 = workbook.getWorksheets().get("Employee");
		  Worksheet sheet1 = workbook.getWorksheets().get("Datasheet1");
		  Worksheet sheet2 = workbook.getWorksheets().get("Datasheet2");


		  System.out.println("demo opened");
		   
		  
		          //Save the sheet to image
		  		System.out.println(sheet0.getIndex());
		          sheet0.saveToImage("D:\\java\\image.png");
		          System.out.println("image converted");
		          
		          
		          File imageFile = new File("D:\\java\\image.png");
		          BufferedImage bufferedImage =null;
		          try
		              {
		              bufferedImage = ImageIO.read(imageFile);
		              BufferedImage image=bufferedImage.getSubimage(0,0,500,500);
		              File pathFile = new File("D:\\java\\image-crop.png");
		              ImageIO.write(image,"jpeg", pathFile);
		            }
		          catch (IOException e) 
		              {
		                  System.out.println(e);
		              }
		  
		      }

	
}
