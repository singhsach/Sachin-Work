package automation_spire;

import java.awt.image.BufferedImage;

import java.io.File;
import java.io.IOException;
import javax.imageio.ImageIO;

public class ExcelToImage {
	public  void excelToImage() throws IOException {
		InputBean input = new InputBean();
		CropImage cropImage = new CropImage();
		
		cropImage.convertExcel(); 
		
		cropImage.cropIt(65, 147, 1020, 364, input.convert(input.daily1),input.convert(input.imagelis));
		cropImage.cropIt(66, 511, 1020, 224, input.convert(input.daily2),input.convert(input.imagelis));
		cropImage.cropIt(63, 730, 1020, 221, input.convert(input.weekly),input.convert(input.imagelis));
		cropImage.cropIt(65, 950, 1020, 150, input.convert(input.monthly),input.convert(input.imagelis));
		cropImage.cropIt(66,70, 775, 485, input.convert(input.rpa1),input.convert(input.imageRpa));
		cropImage.cropIt(65,482, 775, 575, input.convert(input.rpa2), input.convert(input.imageRpa));
		 BufferedImage img1 = ImageIO.read(new File(input.convert(input.weekly)));
		 BufferedImage img2 = ImageIO.read(new File(input.convert(input.monthly)));
		 
		 BufferedImage joinedImg = cropImage.joinBufferedImage(img1, img2);
		 ImageIO.write(joinedImg, "png", new File("D:\\java\\merged.png"));
		 System.out.println("merged file created");
		 
	
		
	}

}
