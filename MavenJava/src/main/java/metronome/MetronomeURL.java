package metronome;

import java.awt.AWTException;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Properties;
import java.util.TimeZone;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.imageio.ImageIO;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class MetronomeURL {

	public static ChromeDriver driver;
	
	@Test(priority = 1)
	public void Metronome_URLlogin() throws IOException, InterruptedException {
		
		
		File src = new File("./ExcelFile/DSC Health Checks.xlsx");
		FileInputStream fis  = new FileInputStream(src);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet1 = wb.getSheet("make changes here");
		
		
		System.setProperty("webdriver.chrome.driver","./driver/chromedriver.exe");
	      ChromeDriver driver = new ChromeDriver();
	      driver.get("https://stepher:131522_JOc@metronome.global.umusic.net/");
	      driver.manage().window().maximize();
	      //driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
	      //Thread.sleep(2000);
	      
	      //ASPEN
	      System.out.println("ASPEN");
	      
	      driver.findElementByXPath("//*[text() = 'Assets & Meta-Data']//following::div[1]//*[text()='ASPEN ']").click();
	      
	      //ASPENServer
	      
	      List<WebElement> ASPENServerOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationServer']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
		     
	      int ASPENServerOKi = ASPENServerOK.size();
	      
	      
	      List<WebElement> ASPENServerERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationServer']//span[text()='ERROR']");
	     
	      int ASPENServerERRORi = ASPENServerERROR.size();
	     
	      
		   if(ASPENServerERRORi > 0){
			   
			   sheet1.getRow(3).createCell(2).setCellValue(ASPENServerERRORi+"/"+ASPENServerOKi+"*");
		   }else {
			  
			   sheet1.getRow(3).createCell(2).setCellValue(ASPENServerOKi);
			   
		   }
		   
	      
		   //ASPEN_Server_Attribute
		   
		   List<WebElement> ASPEN_Server_Attribute = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationServerAttribute']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
		     
		      int ASPEN_Server_Attributei = ASPEN_Server_Attribute.size();
		      
		      
		      List<WebElement> ASPEN_Server_AttributeERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationServerAttribute']//span[text()='ERROR']");
		     
		      int ASPEN_Server_AttributeERRORi = ASPEN_Server_AttributeERROR.size();
		     
		      
			   if(ASPEN_Server_AttributeERRORi > 0){
				   sheet1.getRow(4).createCell(2).setCellValue(ASPEN_Server_AttributeERRORi+"/"+ASPEN_Server_Attributei+"*");
			   }else {
				   sheet1.getRow(4).createCell(2).setCellValue(ASPEN_Server_Attributei);
				   
			   }
			   
			 //Server_Attribute_Disk
			   
			   List<WebElement> Server_Attribute_Disk = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationServerAttributeDisk']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
			     
			      int Server_Attribute_Diski = Server_Attribute_Disk.size();
			      
			      
			      List<WebElement> Server_Attribute_DiskERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationServerAttributeDisk']//span[text()='ERROR' or text() = 'ERROR']");
			     
			      int Server_Attribute_DiskERRORi = Server_Attribute_DiskERROR.size();
			     
			      
				   if(Server_Attribute_DiskERRORi > 0){
					   sheet1.getRow(5).createCell(2).setCellValue(Server_Attribute_DiskERRORi+"/"+Server_Attribute_Diski+"*");
				   }else {
					   sheet1.getRow(5).createCell(2).setCellValue(Server_Attribute_Diski);
					   
				   }
			   
			 
	      
	      //ASPENDatabaseOutput
	      
	      List<WebElement> ASPENDatabaseOutputOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationDatabasePerformance']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
	     
	      int ASPENDatabaseOutputOKi = ASPENDatabaseOutputOK.size();
	      
	      
	      List<WebElement> ASPENDatabaseOutputERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationDatabasePerformance']//span[text()='ERROR']");
	     
	      int ASPENDatabaseOutputERRORi = ASPENDatabaseOutputERROR.size();
	     
	      
		   if(ASPENDatabaseOutputERRORi > 0){
			   sheet1.getRow(7).createCell(2).setCellValue(ASPENDatabaseOutputERRORi+"/"+ASPENDatabaseOutputOKi+"*");
		   }else {
			   sheet1.getRow(7).createCell(2).setCellValue(ASPENDatabaseOutputOKi);
			   
		   }
		   
		   //ASPENInterface
			  
		   List<WebElement> ASPENInterfaceOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationInterface']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
		     
		      int ASPENInterfaceOKi = ASPENInterfaceOK.size();
		      
		      
		      List<WebElement> ASPENInterfaceERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationInterface']//span[text()='ERROR']");
		     
		      int ASPENInterfaceERRORi = ASPENInterfaceERROR.size();
		     
		      
			   if(ASPENInterfaceERRORi > 0){
				   sheet1.getRow(10).createCell(2).setCellValue(ASPENInterfaceERRORi+"/"+ASPENInterfaceOKi+"*");
			   }else {
				   sheet1.getRow(10).createCell(2).setCellValue(ASPENInterfaceOKi);
				   
			   }
			   Thread.sleep(2000);
			   driver.navigate().back();
			   

		        //GPSWPS
				   System.out.println("GPSWPS");
				   
				   driver.findElementByXPath("//*[text() = 'Assets & Meta-Data']//following::div[1]//*[text()='GPS (WS) ']").click();
				   
				   //GPSSERVER
				   
				   List<WebElement> GPSServerOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationServer']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
				     
				      int GPSServerOKi = GPSServerOK.size();
				      
				      
				      List<WebElement> GPSServerERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationServer']//span[text()='ERROR']");
				     
				      int GPSServerERRORi = GPSServerERROR.size();
				     
				      
					   if(GPSServerERRORi > 0){
						   sheet1.getRow(3).createCell(6).setCellValue(GPSServerERRORi+"/"+GPSServerOKi+"*");
					   }else {
						   sheet1.getRow(3).createCell(6).setCellValue(GPSServerOKi);
						   
					   }
					   
					   //GPSDatabaseOutput
					   
					   List<WebElement> GPSDatabaseOutputOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationDatabasePerformance']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
					     
					      int GPSDatabaseOutputOKi = GPSDatabaseOutputOK.size();
					      
					      
					      List<WebElement> GPSDatabaseOutputERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationDatabasePerformance']//span[text()='ERROR']");
					     
					      int GPSDatabaseOutputERRORi = GPSDatabaseOutputERROR.size();
					     
					      
						   if(GPSDatabaseOutputERRORi > 0){
							   sheet1.getRow(7).createCell(6).setCellValue(GPSDatabaseOutputERRORi+"/"+GPSDatabaseOutputOKi+"*");
						   }else {
							   sheet1.getRow(7).createCell(6).setCellValue(GPSDatabaseOutputOKi);
							   
						   }
						   //GPSWindowsService
				   
						   List<WebElement> GPSWindowsServiceOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationWindowsService']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
						     
						      int GPSWindowsServiceOKi = GPSWindowsServiceOK.size();
						      
						      
						      List<WebElement> GPSWindowsServiceERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationWindowsService']//span[text()='ERROR']");
						     
						      int GPSWindowsServiceERRORi = GPSWindowsServiceERROR.size();
						     
						      
							   if(GPSWindowsServiceERRORi > 0){
								   sheet1.getRow(8).createCell(6).setCellValue(GPSWindowsServiceERRORi+"/"+GPSWindowsServiceOKi+"*");
							   }else {
								   sheet1.getRow(8).createCell(6).setCellValue(GPSWindowsServiceOKi);
								   
							   }
							  //GPSInterface
							   
							   List<WebElement> GPSInterfaceOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationInterface']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
							     
							      int GPSInterfaceOKi = GPSInterfaceOK.size();
							      
							      
							      List<WebElement> GPSInterfaceERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationInterface']//span[text()='ERROR']");
							     
							      int GPSInterfaceERRORi = GPSInterfaceERROR.size();
							     
							      
								   if(GPSInterfaceERRORi > 0){
									   sheet1.getRow(10).createCell(6).setCellValue(GPSInterfaceERRORi+"/"+GPSInterfaceOKi+"*");
								   }else {
									   sheet1.getRow(10).createCell(6).setCellValue(GPSInterfaceOKi);
									   
								   }
				   
								   driver.navigate().back();
							   
			   //ASPEN_AGU
			   System.out.println("ASPEN_AGU");
							   
			   driver.findElementByXPath("//*[text() = 'Pricing and Scheduling']//following::div[1]//*[text()='AGU (ASPEN) ']").click();
							   
				//ASPENAGU_URL
							   
     		   List<WebElement> ASPENAGU_URLOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationURL']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
							     
     		      int ASPENAGU_URLOKi = ASPENAGU_URLOK.size();
							      
							      
		      List<WebElement> ASPENAGU_URLERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationURL']//span[text()='ERROR']");
							     
			      int ASPENAGU_URLERRORi = ASPENAGU_URLERROR.size();
							     
							      
		    		   if(ASPENAGU_URLERRORi > 0){
		    			   sheet1.getRow(2).createCell(3).setCellValue(ASPENAGU_URLERRORi+"/"+ASPENAGU_URLOKi+"*");
						   }else {
							   sheet1.getRow(2).createCell(3).setCellValue(ASPENAGU_URLOKi);
									   
							   }
								   
				   //ASPENAGU_Interface
								   
				   List<WebElement> ASPENAGU_InterfaceOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationInterface']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
								     
     			      int ASPENAGU_InterfaceOKi = ASPENAGU_InterfaceOK.size();
								      
								      
				      List<WebElement> ASPENAGU_InterfaceERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationInterface']//span[text()='ERROR']");
								     
				      int ASPENAGU_InterfaceERRORi = ASPENAGU_InterfaceERROR.size();
								     
								      
					   if(ASPENAGU_InterfaceERRORi > 0){
						   sheet1.getRow(10).createCell(3).setCellValue(ASPENAGU_InterfaceERRORi+"/"+ASPENAGU_InterfaceOKi+"*");
				    		   }else {
				    			   sheet1.getRow(10).createCell(3).setCellValue(ASPENAGU_InterfaceOKi);
										   
				        }
			
					   driver.navigate().back();
					   
					   
					   //ASPEN_AST
					   System.out.println("ASPEN_AST");
					   
					   driver.findElementByXPath("//*[text() = 'Pricing and Scheduling']//following::div[1]//*[text()='AST (ASPEN) ']").click();
					   
					   //ASPENAST_URL
					   
					   List<WebElement> ASPENAST_URLOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationURL']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
					     
					      int ASPENAST_URLOKi = ASPENAST_URLOK.size();
					      
					      
					      List<WebElement> ASPENAST_URLERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationURL']//span[text()='ERROR']");
					     
					      int ASPENAST_URLERRORi = ASPENAST_URLERROR.size();
					     
					      
						   if(ASPENAST_URLERRORi > 0){
							   sheet1.getRow(2).createCell(4).setCellValue(ASPENAST_URLERRORi+"/"+ASPENAST_URLOKi+"*");
						   }else {
							   sheet1.getRow(2).createCell(4).setCellValue(ASPENAST_URLOKi);
							   
						   }
						   
		   //ASPENAST_Interface
						   
		   List<WebElement> ASPENAST_InterfaceOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationInterface']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
						     
			      int ASPENAST_InterfaceOKi = ASPENAST_InterfaceOK.size();
						      
						      
		      List<WebElement> ASPENAST_InterfaceERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationInterface']//span[text()='ERROR']");
						     
		      int ASPENAST_InterfaceERRORi = ASPENAST_InterfaceERROR.size();
						     
						      
			   if(ASPENAST_InterfaceERRORi > 0){
				   sheet1.getRow(10).createCell(4).setCellValue(ASPENAST_InterfaceERRORi+"/"+ASPENAST_InterfaceOKi+"*");
		    		   }else {
		    			   sheet1.getRow(10).createCell(4).setCellValue(ASPENAST_InterfaceOKi);
								   
		        }
			   
			   driver.navigate().back();
			   
			   //DIGS
			   System.out.println("DIGS");
			   
			   driver.findElementByXPath("//*[text() = 'Pricing and Scheduling']//following::div[1]//*[text()='DiGS ']").click(); 
			   
			  //DIGS_URL
			   List<WebElement> DIGS_URLOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationURL']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
			     
			      int DIGS_URLOKi = DIGS_URLOK.size();
							      
							      
			      List<WebElement> DIGS_URLERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationURL']//span[text()='ERROR']");
							     
			      int DIGS_URLERRORi = DIGS_URLERROR.size();
							     
							      
				   if(DIGS_URLERRORi > 0){
					   sheet1.getRow(2).createCell(5).setCellValue(DIGS_URLERRORi+"/"+DIGS_URLOKi+"*");
			    		   }else {
			    			   sheet1.getRow(2).createCell(5).setCellValue(DIGS_URLOKi);
									   
			        }
				   //DIGSServer
				   
				   List<WebElement> DIGSServerOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationServer']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
				     
				      int DIGSServerOKi = DIGSServerOK.size();
								      
								      
				      List<WebElement> DIGSServerERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationServer']//span[text()='ERROR']");
								     
				      int DIGSServerERRORi = DIGSServerERROR.size();
								     
								      
					   if(DIGSServerERRORi > 0){
						   sheet1.getRow(3).createCell(5).setCellValue(DIGSServerERRORi+"/"+DIGSServerOKi+"*");
				    		   }else {
				    			   sheet1.getRow(3).createCell(5).setCellValue(DIGSServerOKi);
										   
				        }
					   
					   //DIGSDatabaseOutput
					   
					   List<WebElement> DIGSDatabaseOutputOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationDatabasePerformance']//span[text()='OK' or text() = 'WARNING'or text() = 'ERROR']");
					     
					      int DIGSDatabaseOutputOKi = DIGSDatabaseOutputOK.size();
									      
									      
					      List<WebElement> DIGSDatabaseOutputERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationDatabasePerformance']//span[text()='ERROR']");
									     
					      int DIGSDatabaseOutputERRORi = DIGSDatabaseOutputERROR.size();
									     
									      
						   if(DIGSDatabaseOutputERRORi > 0){
							   sheet1.getRow(7).createCell(5).setCellValue(DIGSDatabaseOutputERRORi+"/"+DIGSDatabaseOutputOKi+"*");
					    		   }else {
					    			   sheet1.getRow(7).createCell(5).setCellValue(DIGSDatabaseOutputOKi);
											   
					        }
						   //DIGSInterface
						   List<WebElement> DIGSInterfaceOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationInterface']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
						     
						      int DIGSInterfaceOKi = DIGSInterfaceOK.size();
										      
										      
						      List<WebElement> DIGSInterfaceERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationInterface']//span[text()='ERROR']");
										     
						      int DIGSInterfaceERRORi = DIGSInterfaceERROR.size();
										     
										      
							   if(DIGSInterfaceERRORi > 0){
								   sheet1.getRow(10).createCell(5).setCellValue(DIGSInterfaceERRORi+"/"+DIGSInterfaceOKi+"*");
						    		   }else {
						    			   sheet1.getRow(10).createCell(5).setCellValue(DIGSInterfaceOKi);
												   
						        }
				        driver.navigate().back();
				        
				        
				        
				    //RWS
				     System.out.println("RWS");
				        
				     driver.findElementByXPath("//*[text() = 'Pricing and Scheduling']//following::div[1]//*[text()='RWS ']").click(); 
				        
				        
				    //RWS_URL
						   
				   List<WebElement> RWS_URLOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationURL']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
						     
			      int RWS_URLOKi = RWS_URLOK.size();
						      
						      
			   List<WebElement> RWS_URLERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationURL']//span[text()='ERROR']");
						     
		      int RWS_URLERRORi = RWS_URLERROR.size();
						     
						      
			   if(RWS_URLERRORi > 0){
								   sheet1.getRow(2).createCell(7).setCellValue(RWS_URLERRORi+"/"+RWS_URLOKi+"*");
							   }else {
								   sheet1.getRow(2).createCell(7).setCellValue(RWS_URLOKi);
								   
								   
							   }
							   
							
							   
			   //RWS_Interface
							   
              
							   
			   List<WebElement> RWS_InterfaceOK = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationInterface']//span[text()='OK' or text() = 'WARNING' or text() = 'ERROR']");
							     
				      int RWS_InterfaceOKi = RWS_InterfaceOK.size();
							      
							      
			      List<WebElement> RWS_InterfaceERROR = driver.findElementsByXPath("//table[@id='MainContent_GVApplicationInterface']//span[text()='ERROR']");
							     
			      int RWS_InterfaceERRORi = RWS_InterfaceERROR.size();
							     
							      
				   if(RWS_InterfaceERRORi > 0){
					   sheet1.getRow(10).createCell(7).setCellValue(RWS_InterfaceERRORi+"/"+RWS_InterfaceOKi+"*");
			    		   }else {
			    			   sheet1.getRow(10).createCell(7).setCellValue(RWS_InterfaceOKi);
									   
			        }
				   System.out.println("Done!!! Please Open the XL File");
				   FileOutputStream fos = new FileOutputStream(src);
					wb.write(fos);
					wb.close();	
					driver.close();
			}

@Test(priority = 2)
	
	public void FilePath()throws AWTException, IOException{
		
		String path = "C:\\Selenium\\Mavenjava\\ExcelFile\\DSC Health Checks.xlsx";
		File file = new File(path);
			//DeskTop Activity
			Robot FileCloseF = new Robot();
	        try{	
	        	
	        	if (file.exists()){
	        		Process pro = Runtime.getRuntime().exec("rundll32 url.dll,FileProtocolHandler "+path);
	        		pro.waitFor();
	        	}else{
	        		System.out.println("File does not exist");
	        	}
	        	FileCloseF .delay(15000);
	        	FileCloseF.keyPress(KeyEvent.VK_CONTROL);
	        	FileCloseF.keyPress(KeyEvent.VK_SHIFT);
	        	FileCloseF.keyPress(KeyEvent.VK_F1);
	        	FileCloseF.keyRelease(KeyEvent.VK_F1);
	        	FileCloseF.keyPress(KeyEvent.VK_F1);
	        	FileCloseF.keyRelease(KeyEvent.VK_F1);
	        	FileCloseF.keyRelease(KeyEvent.VK_SHIFT);
	        	FileCloseF.keyRelease(KeyEvent.VK_CONTROL);
	        	System.out.println("Excel Opened");
	        }catch (Exception e){
	        	System.out.println(e);
	        }
	        FileCloseF .delay(5000);
}
	        
 @Test(priority = 3)
	    	
	    	public void ScreenCap()throws AWTException, IOException{
	        Robot ScreenR = new Robot();
		      ScreenR.delay(1000);
		      Rectangle rectangle = new Rectangle(30,225,920,385);
		         BufferedImage screenShot = ScreenR.createScreenCapture(rectangle);
		         ImageIO.write(screenShot, "png", new File("C:\\Selenium\\MavenJava\\ScreenShots\\DSC_Health_Checks.png"));
		         System.out.println("Screen Saved");
		         ScreenR.delay(1000);
 }
 @Test(priority = 4)     
            public void FileClose()throws AWTException, IOException{
	        Robot FileCloseR = new Robot();
			FileCloseR .delay(15000);
			FileCloseR .keyPress(KeyEvent.VK_ALT);
			FileCloseR .keyPress(KeyEvent.VK_F4);
			FileCloseR .keyRelease(KeyEvent.VK_F4);
			FileCloseR .keyRelease(KeyEvent.VK_ALT);
			System.out.println("Excel Closed");
			FileCloseR .delay(5000);
           }
 @Test(priority = 5)      
	            public void JavaEmail()throws AWTException, IOException{		  
		        final String username = "rstephen@dxc.com";
		 		final String password = "131522_JOc";
		 		String fromEmail = "rstephen@dxc.com";
		 		String toEmail = "ranjith519stephen@gmail.com";
		 		String toEmail1 = "rstephen@dxc.com";
		 		Properties properties = new Properties();
		 		properties.put("mail.smtp.host", "smtp.office365.com");
		 		properties.put("mail.smtp.socketFactory.port", "587"); //SSL Port
		 		properties.put("mail.smtp.starttls.enable", "true");
		 		properties.put("mail.smtp.port", "587");
		 		properties.put("mail.smtp.debug", "true");
		 		properties.put("mail.smtp.auth", "true");
		 		properties.put("mail.smtp.ssl.trust", "smtp.office365.com");
		 		properties.put("mail.smtp.ssl.socketFactory.fallback", false);
		 		

		 		Session session = Session.getInstance(properties, new javax.mail.Authenticator() {
		 			protected PasswordAuthentication getPasswordAuthentication() {
		 				return new PasswordAuthentication(username,password);
		 			}
		 		});
		 		//Start the mail message
		 		MimeMessage msg = new MimeMessage(session);

		 		//Date and Time for the Specific Zone
		 		TimeZone.setDefault(TimeZone.getTimeZone("Canada/Pacific"));

		 		SimpleDateFormat date_format=new SimpleDateFormat("dd/MM/yyyy EEEE hh:mm a");
		 		Date date=new Date();		
		 		String current_date_time=date_format.format(date);		
		 		System.out.println();
		 		try {
		 			msg.setFrom(new InternetAddress(fromEmail));
		 			msg.addRecipient(Message.RecipientType.TO, new InternetAddress(toEmail));
		 			msg.addRecipient(Message.RecipientType.CC, new InternetAddress(toEmail1));
		 			msg.setSubject("DSC Health Check "+current_date_time+" PDT");
		 			MimeMultipart multipart = new MimeMultipart("related");
		 			BodyPart messageBodyPart = new MimeBodyPart();
		 	         String htmlText = "<img src=\"cid:image\">";
		 	         messageBodyPart.setContent(htmlText, "text/html");
		 	         multipart.addBodyPart(messageBodyPart);
		 	         messageBodyPart = new MimeBodyPart();
		 	         DataSource fds = new FileDataSource("C:\\Selenium\\MavenJava\\ScreenShots\\DSC_Health_Checks.png");
		 	         messageBodyPart.setDataHandler(new DataHandler(fds));
		 	         messageBodyPart.setHeader("Content-ID", "<image>");
		 	         multipart.addBodyPart(messageBodyPart);
		 	         msg.setContent(multipart);
		 		    Transport.send(msg);
		 			System.out.println("Sent message");
		 		} catch (AddressException e) {
		 			e.printStackTrace();
		 			
		 	} catch (MessagingException e) {
		 		throw new RuntimeException(e);
		 	}
	}

}
