/**
 * 
 */
package ITimeExport;

import javax.swing.JFrame;
import javax.swing.JPanel;


/**
 * @author 10638193
 *
 */
public class ExportData {

	public static void main(String arg[])
	  {
	    try
	    {
	   	
	    
	    ITimeData extdata=new ITimeData();
        
	    extdata.setSize(1500,750);
	    extdata.setVisible(true);
	    extdata.setTitle("Export Data to Excel");
	    extdata.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	    
	    

      
	  
	    
	    }
	  catch(Exception e)
	    {}
	  }
}
