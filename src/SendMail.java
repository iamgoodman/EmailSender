
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Properties;
import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import javax.mail.PasswordAuthentication;


public class SendMail {


	
	//send email variable
    final String userName ="marketing@grille4u.com";
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
       new SendMail();
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
            Message message = new MimeMessage(session);

            message.setFrom(new InternetAddress("marketing@grille4u.com"));
            message.setRecipients(Message.RecipientType.TO,
                                 /* InternetAddress.parse(emails.get(i).toString()));*/
                                  InternetAddress.parse("joey@apsautoparts.com"));
            message.setSubject("Nice Photo for $15 Cash");
            message.setContent("<h:body>Thank you for purchasing from automaxstyling on ebay<br> We would like to offer you $15 partial refund if you can send us a picture to illustrate our Tonneau Cover on your car.<br> The picture should be taken under bright lighting without flash. Please set your picture size to 2M or or larger, and set the resolution to 180dpi or better. <br></body>","text/html;     charset=utf-8");
            Transport.send(message);
            System.out.println("Done");
            
            
        }catch(MessagingException
               messageException){
            throw new RuntimeException(messageException);
        }

 /*   }*/
    }

}