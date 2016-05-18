
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
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
import javax.swing.plaf.basic.BasicInternalFrameTitlePane.SystemMenuBar;

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
    
    final String password="*****"; 
    
    //arraylist to store emails with emailid , customer id , invoice #
   public static ArrayList emails = new ArrayList();
    
   //arraylist to store invoice # in customers bundle 
   public static ArrayList invoice = new ArrayList<String>();
   


   
   //arraylist to store customerid only in customers bundle 
   public static ArrayList customerid = new ArrayList<String>();
   
                                                                                      
   /*      //arraylist to store emails only in emails bundle for testing
   public static ArrayList emailid = new ArrayList<String>();
   */
   
   
   
    public static void main(String[] args) {
    	

    	//To read input ASINs as excel formatt and parse into java, stored as arrayList
    	
    	 // Location of the source file
       String sourceFilePath = "c:\\email8.xls";
         
       FileInputStream fileInputStream = null;
         
       // Array List to store the excel sheet data
       ArrayList excelData = new ArrayList();
       
       
       

         
       //A more robust importing method for importing excel data to arrays
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
       
     
       
  /*     
       //remove emails we do not want to send.
       for(int i = 0; i< excelData.size(); i++)
       {
           
    	   if(excelData.get(i).toString().contains("****"))
    	   {
    	   excelData.remove(i);
           
    	   }
           
           
       }
   */
    
       //remove the curl brackets in the exceldata arraylist.
       for(int i = 0; i <excelData.size();i++)	
    		
    		
       {
    	   
    	   emails.add(excelData.get(i).toString().substring(1, excelData.get(i).toString().length()-1));
    	   
    	   
    	   
    	 /*System.out.println(emails.get(i));*/
    	   
    	   
    	/*  emailid.add(emails.get(i).toString().substring(0,emails.get(i).toString().indexOf(",")));
    	   
    	   customerid.add( emails.get(i).toString().substring(emails.get(i).toString().indexOf(",")+1));*/
    	   
    	 
    	 //add the remaing part in the string into a array list named invoice
    	   invoice.add(  emails.get(i).toString().split(",") [2]  );
    	   
    	   //get the customer id to customerid list, added to customer id list
    	   customerid.add(emails.get(i).toString().split(",")[1]);
       }
       
       
    /*   //removing trailing 0 s in the invoice string, reduce running time avoided iteration of list
       for (int i = 0; i < invoice.size(); i++)
       {
    	   
    	   
    	   
    	   System.out.println(invoice.get(i).toString().indexOf(".") < 0 ? invoice.get(i).toString() :invoice.get(i).toString().replaceAll("0*$", "").replaceAll("\\.$", "") );
    	   
    	   
       }
       
*/
    	  /* for(int i = 0; i<emailid.size(); i++)
    		   
    	   {
    		   
    		   
    		   
    		   System.out.println(emailid.get(i));
    		   
    		   
    	   }*/
       
/*
	   for(int i = 0; i<customerid.size(); i++)
    		   
    	   {
    		   
    		   
    		   
    		   System.out.println(customerid.get(i));
    		   
    		   
    	   }
       */
       
       
   /*    
       for(int i = 0; i< invoice.size(); i++){
    	   
    	   
    	  System.out.println(invoice.get(i));
    	   
    	   
       }*/
    	   
    
       

       new SendMail(emails);
       
   
       
    /* String emailtitlemessage = String.format("Nice Photo for $15 Partial Refund Ebay Invoice # %.0f\n", 
    		   Double.parseDouble(invoice.get(0).toString()));
       
       
       System.out.println(emailtitlemessage);*/
     
       
     
    }

    public SendMail(ArrayList emailss){
    	
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
        
        //send  a list of emails in the arraylist emails

      for(int i = 0; i< emails.size();i++)
    	  
        {
    	  
    
    		  
    	  
        try{
        	
        	
        	
        	//send email
        	
        	
            Message message = new MimeMessage(session);

            message.setFrom(new InternetAddress("marketing@grille4u.com"));
            message.setRecipients(Message.RecipientType.TO,
                 
                 InternetAddress.parse(emails.get(i).toString().substring(0,emails.get(i).toString().indexOf(","))));
       
            
     
            
            
            //set title message, remove the scientific notation
            
            
            String emailtitlemessage = String.format("Nice Photo for $15 Partial Refund Ebay Invoice # %.0f\n", 
         		   Double.parseDouble(invoice.get(i).toString()));
            
            
            
            System.out.println(emailtitlemessage + "This message is ready to be send ");
            
            
            message.setSubject(emailtitlemessage);
           
         
            
            //set message body of email, make a part of message bold and red
            String bodymessageinput = String.format("<h:body>Dear %s <br><br><br>Thank you for purchasing from automaxstyling on ebay.<br><br> We would like to offer you $15 partial refund if you can send us a picture to illustrate our Tonneau Cover on your car.<br><br> The picture should be taken under bright lighting without flash. Please set your picture size to 2M or or larger, and set the resolution to 180dpi or better.<br><br><b><i><font color=\"red\"> Please follow the sample pictures in the attachment below, and reply this email to send your photo to us</font></i></b><br><br>Thank you very much!"
            		+ "<br><br><br>AutomaxStyling From Ebay</body>", customerid.get(i) );
     
           /* emails.get(i).toString().substring(emails.get(i).toString().indexOf(",")+1*/
            
            //set boy part of the email
            MimeBodyPart messageBodyPart = new MimeBodyPart();
            
      
            
        	// Fill the message
			messageBodyPart.setText(bodymessageinput);
			messageBodyPart.setContent(bodymessageinput, "text/html");
            
            
            
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
            
          
            
            System.out.println( message+"has been Send");
            
            
        }catch(MessagingException
        		
               messageException){
        	
        	
            throw new RuntimeException(messageException);
                            
        }

    }
      
      
      System.out.println("all emails send ");
      
        
        }
        
    

    
    
    
}