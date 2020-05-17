/**
 * 
 */
package ITimeExport;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * @author 10638193
 *
 */
public class PDFtoConsole {

	public static void main(String args[]) throws IOException {
        FileInputStream fis = new FileInputStream(new File("C:\\Excel toPDF\\Excel2PDF_Output.pdf"));
        ByteArrayOutputStream docContents = new ByteArrayOutputStream();
        byte[] buffer = new byte[16384];
        int bytesRead = fis.read(buffer);
        while (bytesRead > -1) {
            docContents.write(buffer, 0, bytesRead);            
            bytesRead = fis.read(buffer);
            
        }
        System.out.println(docContents.toString("UTF-8"));
    }
 
	}


