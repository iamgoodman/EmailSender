
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Properties;


import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.nio.DataSource;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import javax.mail.PasswordAuthentication;


public class SendMail {


	
	//send email variable
    final String userName ="marketing@grille4u.com";
    //np
    final String password="***"; 
   public static ArrayList emails = new ArrayList();
    
    
    public static void main(String[] args) {
    	

    	//To read input ASINs as excel formatt and parse into java, stored as arrayList
    	
    	 // Location of the source file
       String sourceFilePath = "c:\\email1.xls";
         
       FileInputStream fileInputStream = null;
         
       // Array List to store the excel sheet data
       ArrayList excelData = new ArrayList();
         
       try {
             
           // FileInputStream to read the excel file
           fileInputStream = new FileInputStream(sourceFilePath);

           // Create an excel workbook
           HSSFWorkbook excelWorkBook = new HSSFWorkbook(fileInputStream);
             
           // Retrieve the first sheet of the workbook.
           HSSFSheet excelSheet = excelWorkBook.getSheetAt(0);

           // Iterate through the sheet rows and cells. 
           // Store the retrieved data in an arrayList
           java.util.Iterator<Row> rows = excelSheet.rowIterator();
           while (rows.hasNext()) {
               HSSFRow row = (HSSFRow) rows.next();
               java.util.Iterator<Cell> cells = row.cellIterator();

               ArrayList cellData = new ArrayList();
               while (cells.hasNext()) {
                   HSSFCell cell = (HSSFCell) cells.next();
                   cellData.add(cell);
               }

               excelData .add(cellData);
           }
             
           // Print retrieved data to the console
           for (int rowNum = 0; rowNum < excelData.size(); rowNum++) {
                 
               ArrayList list = (ArrayList) excelData.get(rowNum);
                 
               for (int cellNum = 0; cellNum < list.size(); cellNum++) {
                     
                   HSSFCell cell = (HSSFCell) list.get(cellNum);
                     
                 
               }
               
           }
       } catch (IOException e) {
           e.printStackTrace();
       } finally {
           if (fileInputStream != null) {
               try {
				fileInputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
           }
       }
       for(int i = 0; i <excelData.size();i++)	
    		
    		
       {
    	   
    	   emails.add(excelData.get(i).toString().substring(1, excelData.get(i).toString().length()-1));
    	   
       }
 
     /*   new SendMail(emails);*/
       
       
      /* new SendMail();*/
    }

    public SendMail(/*ArrayList l*/){
    	
        Properties properties = new Properties();
        properties.put("mail.smtp.host", "smtp.gmail.com");
        properties.put("mail.smtp.port", "587");
        properties.put("mail.smtp.starttls.enable", "true") ;
        properties.put("mail.smtp.auth", "true") ;

        Session session = Session.getInstance(properties,new Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication(){
                return new PasswordAuthentication(userName, password);
            }

        });

/*      for(int i = 0; i< emails.size();i++)
        {*/
        
        try{
        	
        	
        	
        	//send email
        	
        	
            Message message = new MimeMessage(session);

            message.setFrom(new InternetAddress("marketing@grille4u.com"));
            message.setRecipients(Message.RecipientType.TO,
                                
                                  InternetAddress.parse("joey@apsautoparts.com"));
          /*  message.setSubject("Nice Photo for $15 Cash");*/
            /*message.setContent("<h:body>Thank you for purchasing from automaxstyling on ebay<br> We would like to offer you $15 partial refund if you can send us a picture to illustrate our Tonneau Cover on your car.<br> The picture should be taken under bright lighting without flash. Please set your picture size to 2M or or larger, and set the resolution to 180dpi or better. <br></body>","text/html;     charset=utf-8");*/
            
            
            message.setSubject("Photo for $15 Partial Refund!! ");
           
            
            //set message body of email
            String s = "<h:body>Dear Customer:<br><br><br>Thank you for purchasing from automaxstyling on ebay.<br><br> We would like to offer you $15 partial refund if you can send us a picture to illustrate our Tonneau Cover on your car.<br><br> The picture should be taken under bright lighting without flash. Please set your picture size to 2M or or larger, and set the resolution to 180dpi or better.<br><br> Please follow the sample pictures below, and reply via this email to send your photo to us<br><br>Thank you very much!<br><br><br>AutomaxStyling From Ebay</body>";
     
          
            
            //set boy part of the email
            MimeBodyPart messageBodyPart = new MimeBodyPart();
            
      
            
        	// Fill the message
			messageBodyPart.setText(s);
			messageBodyPart.setContent(s, "text/html");
            
            
            
	        //set attachment part of the email
            MimeBodyPart attachmentpart = new MimeBodyPart();
            
            
            
            String file = "c:\\expic.jpg";
            String fileName = "attachmentName";
            FileDataSource source = new FileDataSource(file);
            
            
            attachmentpart.setDataHandler(new DataHandler(source));
            attachmentpart.setFileName(fileName);
        
            
            //use multipart to append both message and attachment
			
            Multipart multipart = new MimeMultipart();
            

        
            
            multipart.addBodyPart(attachmentpart);
           
            multipart.addBodyPart(messageBodyPart);
            
            

            message.setContent(multipart);
            
            
       
            
            
            Transport.send(message);
            System.out.println("Done");
            
            
        }catch(MessagingException
        		
               messageException){
        	
        	
            throw new RuntimeException(messageException);
            
            
            
        }

 /*   }*/
        
        
        
    }

    
    
    
}