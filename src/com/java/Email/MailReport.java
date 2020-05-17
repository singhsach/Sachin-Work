package com.java.Email;
import java.io.FileInputStream;
import java.time.Month;

import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.activation.DataHandler;

import javax.activation.DataSource;

import javax.activation.FileDataSource;

import javax.mail.*;

import javax.mail.internet.InternetAddress;

import javax.mail.internet.MimeBodyPart;

import javax.mail.internet.MimeMessage;

import javax.mail.internet.MimeMultipart;

import ITimeExport.ReadExcel;
import automation_spire.InputBean;

public class MailReport {

	
	public void sendMail() throws Exception

	{
//		ReadExcel re=new ReadExcel();
//		int count =re.readExcel();
		FileInputStream input = new FileInputStream("C:\\Users\\10638193\\workspace\\HackathonProject\\resources\\config.properties");
		Properties config = new Properties ();
		config.load(input);
		InputBean bean=new InputBean();

		String host = config.getProperty("host");//host name
		String from = config.getProperty("uname");//sender id
		String to = config.getProperty("reciever");//reciever id
		String pass =config.getProperty("pwd");//sender's password 
		String fileAttachment =config.getProperty("fileAttachment")+bean.getDate()+".pptx";//file name for attachment 
        String port=config.getProperty("port");
        
       System.out.println(pass);
		//system properties

		Properties prop = System.getProperties();

		// Setup mail server properties

	    prop.put("mail.smtp.host", host);
		prop.put("mail.smtp.starttls.enable", "true");
		prop.put("mail.smtp.user", from);
		prop.put("mail.smtp.password", pass);
		prop.put("mail.smtp.port", port);
		prop.put("mail.smtp.auth", "true");

		//session 

		Session session = Session.getInstance(prop, null);

		// Define message
		try {

//			TimesheetMonth t=new TimesheetMonth();
//			Month month=t.getCurrentMonth();
			

			MimeMessage message = new MimeMessage(session);
			message.setFrom(new InternetAddress(from));
			message.addRecipient(Message.RecipientType.TO,new InternetAddress(to));
			message.setSubject("Infoscreen Report "+ bean.getDate());

			// create the message part 
			String[] name = from.split("\\.");

			String[] surname = name[1].split("@");

			MimeBodyPart messageBodyPart = new MimeBodyPart();
			
			messageBodyPart.setText("Hi All,\n\nPlease find below Info Screen Report.\n\nThanks & Regards, \n" + name[0] +" " + surname[0]);
//
//			if(count==0)
//				messageBodyPart.setText("Hi, \n No Mismatch Found.\n\nThanks & Regards, \n" + name[0] +" " + surname[0]);
//			//message body
//			else
//				messageBodyPart.setText("Hi,\n\n Total mismatches: "+count+"\n\nThanks & Regards, \n" + name[0] +" " + surname[0]);

			//messageBodyPart.setText("***");
			Multipart multipart = new MimeMultipart();
			multipart.addBodyPart(messageBodyPart);

			//attachment

			messageBodyPart = new MimeBodyPart();
			DataSource source = new FileDataSource(config.getProperty("filePath")+bean.getDate()+".pptx");
			messageBodyPart.setDataHandler(new DataHandler(source));
			messageBodyPart.setFileName(fileAttachment);
			multipart.addBodyPart(messageBodyPart);
			message.setContent(multipart);

			//send message to reciever

			Transport transport = session.getTransport("smtp");
	        transport.connect(host, from, pass);
			transport.sendMessage(message, message.getAllRecipients());
			System.out.println("Mail Sent Successfully!");
			transport.close();

		}


		catch (AuthenticationFailedException e){
			java.util.logging.Logger.getLogger("Test").log(Level.INFO, "BOOM!", e);
		}
	}
}

