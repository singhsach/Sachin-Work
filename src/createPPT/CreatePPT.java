/**
 * 
 */
package createPPT;

import java.awt.Color;
import java.awt.Rectangle;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 * @author 10638193
 *
 */
public class CreatePPT {

	/**
	 * @param args
	 */
	public static void main(String[] args) {

		XMLSlideShow ppt = new XMLSlideShow();  
		XSLFSlideMaster master = ppt.getSlideMasters().get(0);

		XSLFSlideLayout layout= master.getLayout(SlideLayout.TITLE_ONLY);
		

		try {     
			XSLFSlide slide = ppt.createSlide(layout); 
			XSLFTextShape title= slide.getPlaceholder(0);

			title.clearText();
			XSLFTextParagraph p= title.addNewTextParagraph();
			XSLFTextRun r= p.addNewTextRun();
			r.setText("Info Screen");
			r.setFontColor(Color.black);
			r.setFontSize(20.);

			java.awt.Dimension pgsize = ppt.getPageSize();  
			int width = pgsize.width;  //slide width in points  
			int height = pgsize.height; //slide height in points  
			System.out.println("width: "+  width);  
			System.out.println("height: "+ height); 
			byte[] pictureData = IOUtils.toByteArray(new FileInputStream("D:\\Image\\img.png"));  
            XSLFPictureData pd = ppt.addPicture(pictureData, XSLFPictureData.PictureType.PNG);  
			XSLFPictureShape pic = slide.createPicture(pd);

			pic.setAnchor(new Rectangle(10,75,700,420));

			FileOutputStream out = new FileOutputStream("D:\\PPTX\\InfoPPT.pptx");
			ppt.write(out);
			out.close();

			//ppt.write(os);  
			System.out.println("Image has been written succesfully");
		}
		catch(Exception e) {  
			System.out.println(e);  
		}  
	}

}
