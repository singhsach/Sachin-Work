
package createPPT;


import java.io.FileOutputStream;
import java.io.FileInputStream;

import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.sl.usermodel.*;

import org.openxmlformats.schemas.presentationml.x2006.main.*;
import org.openxmlformats.schemas.drawingml.x2006.main.*;

import java.awt.Dimension;

public class PPTTemplate {

 public static void main(String[] args) throws Exception {

  XMLSlideShow slideShow = new XMLSlideShow();
  XSLFPictureData[] pictureDatas = new XSLFPictureData[]{
   slideShow.addPicture(new FileInputStream("D:\\Capture.png"), PictureData.PictureType.PNG),
   slideShow.addPicture(new FileInputStream("D:\\Capture1.png"), PictureData.PictureType.PNG)};

  // slideShow.addPicture(new FileInputStream("Desert.jpg"), PictureData.PictureType.PNG),
  // slideShow.addPicture(new FileInputStream("Chrysanthemum.jpg"), PictureData.PictureType.PNG)};
//s slides, each having one different background picture out of pictureDatas array
 for (int s = 0; s < pictureDatas.length; s++ ) {
  XSLFSlide slide = slideShow.createSlide();
  CTBackgroundProperties backgroundProperties = slide.getXmlObject().getCSld().addNewBg().addNewBgPr();
  CTBlipFillProperties blipFillProperties = backgroundProperties.addNewBlipFill();
  CTRelativeRect ctRelativeRect = blipFillProperties.addNewStretch().addNewFillRect();
  String idx = slide.addRelation(null, XSLFRelation.IMAGES, pictureDatas[s]).getRelationship().getId();
  CTBlip blib = blipFillProperties.addNewBlip();
  blib.setEmbed(idx);
 }
 
 

 FileOutputStream out = new FileOutputStream("D:\\CreatePPTXSheetsDifferentBackgroundPictures.pptx");
 slideShow.write(out);
 out.close();
}
}