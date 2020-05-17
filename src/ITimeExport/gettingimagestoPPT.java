/**
 * 
 */
package ITimeExport;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import com.aspose.slides.IOleObjectFrame;
import com.aspose.slides.ISlide;
import com.aspose.slides.License;
import com.aspose.slides.Presentation;
import com.aspose.slides.SaveFormat;

/**
 * @author 10638193
 *
 */
public class gettingimagestoPPT {
	
		

	

		public  void setAsposeLicense() {
			InputStream inputStream = null;
			try {
				License license = new License();
				inputStream = getClass().getClassLoader().getResourceAsStream("Aspose_Total.lic");
				license.setLicense(inputStream);
			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				try {
					if (inputStream != null)
						inputStream.close();
				} catch (Exception e) {
					inputStream = null;
				}
			}
		}
		public static void main(String[] args) throws IOException {
			   gettingimagestoPPT ppt=new gettingimagestoPPT();
				ppt.setAsposeLicense();
				//Instantiate Prseetation class that represents the PPTX
				Presentation pres = new Presentation();
				//Access the first slide
				ISlide sld = pres.getSlides().get_Item(0);
				//Load an Excel file to Array of Bytes
				File file=new File("C:\\Itime Sheet\\ITimeData.xlsx");
				int length=(int)file.length();
				FileInputStream fstro = new FileInputStream(file);
				byte[] buf = new byte[length];
				fstro.read(buf, 0, length);
				//Add an Ole Object Frame shape
				IOleObjectFrame oof = sld.getShapes().addOleObjectFrame((float)0,(float) 0, (float)pres.getSlideSize().getSize().getWidth(),(float) pres.getSlideSize().getSize().getHeight (), "Excel.Sheet.8", buf);
				//Write the PPTX to disk
				pres.save("D:\\PPTX\\Reporting.pptx", SaveFormat.Pptx);
				System.out.println("ppt generated.........");

			}
			 		
		}
	


