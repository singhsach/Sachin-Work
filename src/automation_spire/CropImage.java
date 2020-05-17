package automation_spire;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.imageio.ImageIO;

import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;

public class CropImage {
	public void convertExcel() {

		InputBean input = new InputBean();
		Workbook workbook = new Workbook();

		workbook.loadFromFile(input.getExcel());

		// Get the first worksheet
		Worksheet sheet = workbook.getWorksheets().get("LIS-DW-LOAN");

		// Save the sheet to image
		sheet.saveToImage(input.convert(input.imagelis));
		System.out.println("image lis dw loan converted");

		sheet = workbook.getWorksheets().get("RPA");
		sheet.saveToImage(input.convert(input.imageRpa));
		System.out.println("image rpa converted");
	}

	public void cropIt(int x, int y, int width, int height, String file, String ipfile) {
		File imageFile = new File(ipfile);
		BufferedImage bufferedImage = null;
		try {
			bufferedImage = ImageIO.read(imageFile);

			// 63, 72, 937, 491
			BufferedImage image = bufferedImage.getSubimage(x, y, width, height);

			// "D:\\java\\daily1.png"
			File pathFile = new File(file);
			ImageIO.write(image, "png", pathFile);
			System.out.println("cropeed image stored at" + file);
		} catch (IOException e) {
			System.out.println(e);
		}
	}

	public BufferedImage joinBufferedImage(BufferedImage img1, BufferedImage img2) {
		int offset = 2;
		int width = Math.max(img1.getWidth(), img2.getWidth()) + offset;
		int height = img1.getHeight() + img2.getHeight() + offset;
		BufferedImage newImage = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
		Graphics2D g2 = newImage.createGraphics();
		Color oldColor = g2.getColor();
		g2.setPaint(Color.BLACK);
		g2.fillRect(0, 0, width, height);
		g2.setColor(oldColor);
		g2.drawImage(img1, null, 0, 0);
		g2.drawImage(img2, null, 0, img1.getHeight());
		g2.dispose();
		return newImage;
	}
}
