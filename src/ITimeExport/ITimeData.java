/**
 * 
 */
package ITimeExport;

import java.awt.*;
import java.io.*;
import java.awt.event.*;
import java.sql.*;
import java.text.DateFormat;
//import java.util.Date;
import java.util.*;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextArea;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hpsf.Date;
import org.apache.poi.hssf.usermodel.HSSFCell;

/**
 * @author 10638193
 *
 */
public class ITimeData extends JFrame{
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	ResultSet rs;
    ITimeData(){
    final Vector columnNames = new Vector();
        final Vector data = new Vector();
        JButton button=new JButton("Export Data");
    JPanel panel=new JPanel();
    JPanel panel1=new JPanel();
    JPanel panel2=new JPanel();
    JPanel title = new JPanel();
    
    JLabel header = new JLabel("<html><span style='color: teal;'>Export Data</span></html>");
    header.setFont(header.getFont().deriveFont(20.0F));
    
      try{
    Connection con = null;
    Class.forName("com.mysql.cj.jdbc.Driver");
      con = DriverManager.getConnection("jdbc:mysql://sql12.freesqldatabase.com", "sql12331868", "3Fk7JMaxjW");
        Statement st = con.createStatement();
        rs= st.executeQuery("SELECT * FROM sql12331868.Employee");
    ResultSetMetaData md = rs.getMetaData();
int columns = md.getColumnCount();
for (int i = 1; i <= columns; i++) {
columnNames.addElement( md.getColumnName(i) );
}
while (rs.next()) {
Vector row = new Vector(columns);
for (int i = 1; i <= columns; i++) {
row.addElement( rs.getObject(i) );
}
data.addElement(row);
}
}
catch(Exception e){}
JTable table = new JTable(data, columnNames);
JScrollPane scrollPane = new JScrollPane(table);
panel1.add(scrollPane);
panel2.add(button);
panel.add(panel1);
panel.add(panel2);
title.add(header);
add(panel);
button.addActionListener(new ActionListener(){
public void actionPerformed(ActionEvent ev){

	
        String jdbcURL = "jdbc:mysql://sql12.freesqldatabase.com";
        String username = "sql12331868";
        String password = "3Fk7JMaxjW";
 
        String excelFilePath = "C:\\Itime Sheet\\ITimeData.xlsx";
        
        String str = ev.getActionCommand();
  	  if(str.equals("Export Data")){
  	  JOptionPane.showMessageDialog(null, "Data has been exported successfully", "Export Data", 1);
  	  }
  	  
  	
  	  
        try (Connection connection = DriverManager.getConnection(jdbcURL, username, password)) {
            String sql = "SELECT * FROM sql12331868.Employee";
 
            Statement statement = connection.createStatement();
 
            ResultSet result = statement.executeQuery(sql);
 
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Employee");
 
            writeHeaderLine(sheet);
 
            writeDataLines(result, workbook, sheet);
 
            FileOutputStream outputStream = new FileOutputStream(excelFilePath);
            workbook.write(outputStream);
            workbook.close();
 
            statement.close();
            System.out.println("Data has been fetched successfully");
 
        } catch (SQLException e) {
            System.out.println("Datababse error:");
            e.printStackTrace();
        } catch (IOException e) {
            System.out.println("File IO error:");
            e.printStackTrace();
        }
    }
 
    private void writeHeaderLine(XSSFSheet sheet) {
 
        Row headerRow = sheet.createRow(0);
 
        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellValue("PSNumber");
 
        headerCell = headerRow.createCell(1);
        headerCell.setCellValue("Name");
 
        headerCell = headerRow.createCell(2);
        headerCell.setCellValue("TimeBookingDate");
 
        headerCell = headerRow.createCell(3);
        headerCell.setCellValue("BillableHrs");
 
        headerCell = headerRow.createCell(4);
        headerCell.setCellValue("NonBillableHrs");
        
        headerCell = headerRow.createCell(5);
        headerCell.setCellValue("OverTimeHrs");
        
        headerCell = headerRow.createCell(6);
        headerCell.setCellValue("HolidayWorkingHrs");
        
        headerCell = headerRow.createCell(7);
        headerCell.setCellValue("Status");
    }
 
    private void writeDataLines(ResultSet result, XSSFWorkbook workbook,
            XSSFSheet sheet) throws SQLException {
        int rowCount = 1;
 
        while (result.next()) {
            Integer psnumber = result.getInt("PSNumber");
            String name = result.getString("Name");
            java.util.Date timesheetbookingdate = result.getDate("TimeBookingDate");
            Integer billabelhrs = result.getInt("BillableHrs");
            Integer nonbillabelhrs = result.getInt("NonBillableHrs");
            Integer overtimehrs = result.getInt("OverTimeHrs");
            Integer holidayworkinghrs = result.getInt("HolidayWorkingHrs");
            String status = result.getString("Status");
 
            Row row = sheet.createRow(rowCount++);
 
            int columnCount = 0;
            Cell cell = row.createCell(columnCount++);
            cell.setCellValue(psnumber);
 
            cell = row.createCell(columnCount++);
            cell.setCellValue(name);
 
            CellStyle cellStyle = workbook.createCellStyle();
    		CreationHelper createHelper = workbook.getCreationHelper();
    		short dateFormat = createHelper.createDataFormat().getFormat("M/D/YY");
    		cellStyle.setDataFormat(dateFormat);
    		
            cell = row.createCell(columnCount++);
            cell.setCellValue(timesheetbookingdate);
            cell.setCellStyle(cellStyle);
 
            cell = row.createCell(columnCount++);
            cell.setCellValue(billabelhrs+overtimehrs+holidayworkinghrs);
            
            cell = row.createCell(columnCount++);
            cell.setCellValue(nonbillabelhrs);
            
            cell = row.createCell(columnCount++);
            cell.setCellValue(overtimehrs);
            
            cell = row.createCell(columnCount++);
            cell.setCellValue(holidayworkinghrs);
            
            cell = row.createCell(columnCount++);
            cell.setCellValue(status);
 
        }

}
	
	
}

);
}

	
}
