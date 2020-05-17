package FileCount;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.RandomAccessFile;


public class FileCount {


	public static void main(String[] args) {
		FileCount rf = new FileCount();
		  File file = new File("G:\\Automation.txt");
		  rf.readFromLast(file);
		 }
		 
		 public void readFromLast(File file){
		  
		        int lines = 0;
		        StringBuilder builder = new StringBuilder();
		        RandomAccessFile randomAccessFile = null;
		        try {
		         randomAccessFile = new RandomAccessFile(file, "r");
		         long fileLength = file.length() - 1;
		         randomAccessFile.seek(fileLength);
		         for(long pointer = fileLength; pointer >= 0; pointer--){
		           randomAccessFile.seek(pointer);
		           char c;
		           c = (char)randomAccessFile.read(); 
		           if(c == '\n'){
		              break;
		           }
		           builder.append(c);
		         }
		         
		         builder.reverse();
		         System.out.println("No. Of Records - " + builder.toString().substring(8,12));
		         
		        } catch (FileNotFoundException e) {
		           e.printStackTrace();
		        }
		        catch (IOException e) {
		          e.printStackTrace();
		        }finally{
		           if(randomAccessFile != null){
		              try {
		                 randomAccessFile.close();
		              } catch (IOException e) {
		                 e.printStackTrace();
		              }
		           }
		       }
		  }


}
