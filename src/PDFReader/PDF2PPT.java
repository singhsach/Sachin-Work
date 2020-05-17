/**
 * 
 */
package PDFReader;


import java.awt.Color;
import java.awt.Rectangle;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFRelation;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlip;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlipFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRelativeRect;
import org.openxmlformats.schemas.presentationml.x2006.main.CTBackgroundProperties;

import automation_spire.InputBean;




/**
 * @author 10638193
 *
 */
public class PDF2PPT {
	

	public  void convertPDF2Image(String foldername) throws Exception {

			File f = new File(foldername);
			File[] allSubFiles = f.listFiles();
			for(File file : allSubFiles){
				if(file.isDirectory()){
				System.out.println(file.getAbsolutePath()+" is directory");	
				}
				else
				{
					file=new File(file.getAbsolutePath());
					System.out.println(file.getAbsolutePath());
					
   
			        if (file.exists()) {
			            PDDocument document = PDDocument.load(file);
			            String[] fileName =file.getName().split(".pdf");
			            PDPage page=document.getPage(0);
			            page.setCropBox(new PDRectangle(60,410,480,285));
			            PDFRenderer renderer =new PDFRenderer(document);
			            BufferedImage image = renderer.renderImageWithDPI(0,150);
			            File fileImage = new File (Constant.IMAGE_ROOT_FILE_PATH);
			            if(!fileImage.exists()) {
			            	fileImage.mkdir();
			            }
			            
			            ImageIO.write(image,"PNG",new File(Constant.IMAGE_ROOT_FILE_PATH+fileName[0]+".PNG"));
			            
			            document.close();
						System.out.println("Image has been taken Successfully");

			         }   else {
			        	 continue;
			         }
		         	
				}
				
			}
	}	

	
	
	public  void convertPDF2PPT(String imageFolderName) throws Exception {
		System.out.println("1");
		//creating presentation
		XMLSlideShow ppt=new XMLSlideShow();
		System.out.println(imageFolderName);
		File f =new File(imageFolderName);
		File[] allSubFiles=f.listFiles();
		for(File file : allSubFiles) {
			if(!file.isDirectory())
			{   //reading an image
				System.out.println(file.getAbsolutePath());
				file=new File(file.getAbsolutePath());
				  
				  
				XSLFSlideMaster master = ppt.getSlideMasters().get(0);

				XSLFSlideLayout layout= master.getLayout(SlideLayout.TITLE_ONLY);
				
				 

				//creating slide
				org.apache.poi.xslf.usermodel.XSLFSlide slide=ppt.createSlide(layout);
				XSLFTextShape title= slide.getPlaceholder(0);

				title.clearText();
				XSLFTextParagraph p= title.addNewTextParagraph();
				XSLFTextRun r= p.addNewTextRun();
				r.setText(file.getName());
				r.setFontColor(Color.black);
				r.setFontSize(20.);
				
				//converting it into a byte array
				byte[] picture =IOUtils.toByteArray(new FileInputStream(file));
				
				//adding the image to the ppt
				XSLFPictureData idx=ppt.addPicture(picture,XSLFPictureData.PictureType.PNG);
				
				
				
                 
				//creating a slide with given picture
				XSLFPictureShape pic=slide.createPicture(idx);
				
				
					
				
				pic.setAnchor(new Rectangle(10,75,700,420));
				InputBean bean=new InputBean();
				//creating a file object
				File file1=new File("D:/PPTX/Info Screen Report"+bean.getDate()+".pptx");
				FileOutputStream out=new FileOutputStream(file1);
				
				//saving the changes to file
				ppt.write(out);
				System.out.println("Image has been written to PPT Successfully");
				//file.delete();
				file.deleteOnExit();
				System.out.println("image delete");
			}
			
		}
		
		
	}

	
	
	
	


}
