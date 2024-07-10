package com.bdo.Apache;

import java.awt.BorderLayout;


import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;
 
import javax.mail.Authenticator;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.SwingUtilities;
 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import net.sf.json.JSONArray;
	    import org.json.JSONArray;
	    //import org.json.JSONObject;
import org.json.simple.JSONObject;
 
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
//import com.jacob.activeX.ActiveXComponent;

//import com.jacob.activeX.ActiveXComponent;
//import com.jacob.com.ComThread;
//import com.jacob.com.Dispatch;

//import com.jacob.com.ComThread;
//import com.jacob.com.Dispatch;

public class ReadGUI extends JFrame {
	
	 private JButton selectEmailFileButton;
	   private JButton selectDataFileButton;
	   private JButton sendEmailButton;
	   private JLabel statusLabel;
	   private JProgressBar progressBar;
	   private File emailFile;
	   private File dataFile;

	    public ReadGUI() {
	        initializeUI();
	    }

	    private void initializeUI() {
	        setTitle("Email Sender Application");
	        setSize(400, 250);
	        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

	        // Create buttons
	        selectEmailFileButton = new JButton("Select Email File");
	        selectDataFileButton = new JButton("Select Data File");
	        sendEmailButton = new JButton("Send Email");

	        // Create status label
	        statusLabel = new JLabel("Select email and data files to start.....");
	        progressBar = new JProgressBar(0, 100);
	        progressBar.setSize(50, 25);
	        progressBar.setStringPainted(true);
	        setLayout(new BorderLayout());
	        
	        add(selectEmailFileButton);
	        add(selectDataFileButton);
	        add(sendEmailButton);
	        add(statusLabel);
	        add(progressBar);
	        JPanel buttonPanel = new JPanel();
	        buttonPanel.setLayout(new FlowLayout());
	        buttonPanel.add(selectEmailFileButton);
	        buttonPanel.add(selectDataFileButton);
	        buttonPanel.add(sendEmailButton);

	        // Add components to the frame
	        add(buttonPanel, BorderLayout.NORTH);
	        add(progressBar, BorderLayout.CENTER);
	        add(statusLabel, BorderLayout.SOUTH);

	        // Button actions
	        selectEmailFileButton.addActionListener(new ActionListener() {
	            @Override
	            public void actionPerformed(ActionEvent e) {
	                selectEmailFile();
	            }
	        });

	        selectDataFileButton.addActionListener(new ActionListener() {
	            @Override
	            public void actionPerformed(ActionEvent e) {
	                selectDataFile();
	            }
	        });

	        sendEmailButton.addActionListener(new ActionListener() {
	            @Override
	            public void actionPerformed(ActionEvent e) {
	                processAndSendEmails();
	            }
	        });

	        setVisible(true);
	    }

	    private void selectEmailFile() {
	        JFileChooser fileChooser = new JFileChooser();
	        int option = fileChooser.showOpenDialog(this);
	        if (option == JFileChooser.APPROVE_OPTION) {
	            emailFile = fileChooser.getSelectedFile();
	            statusLabel.setText("Email file selected: " + emailFile.getAbsolutePath());
	        } else {
	            statusLabel.setText("No email file selected.");
	        }
	    }

	    private void selectDataFile() {
	        JFileChooser fileChooser = new JFileChooser();
	        int option = fileChooser.showOpenDialog(this);
	        if (option == JFileChooser.APPROVE_OPTION) {
	            dataFile = fileChooser.getSelectedFile();
	            statusLabel.setText("Data file selected: " + dataFile.getAbsolutePath());
	        } else {
	            statusLabel.setText("No data file selected.");
	        }
	    }

	    private void processAndSendEmails() {
	        if (emailFile == null || dataFile == null) {
	            JOptionPane.showMessageDialog(this, "Please select email and data files first.");
	            return;
	        }

	        // Call your existing logic to process and send emails
	        try {
	            // Perform operations with emailFile and dataFile paths
	            String emailFilePath = emailFile.getAbsolutePath();
	            String dataFilePath = dataFile.getAbsolutePath();

	            // Example call to your existing logic
	            processEmails(emailFilePath, dataFilePath);

	            // Optionally update status
	            progressBar.setValue(100);
	            statusLabel.setText("Emails sent successfully.");
	        } catch (Exception e) {
	            e.printStackTrace();
	            statusLabel.setText("Error sending emails: " + e.getMessage());
	        }
	    }

	    public static void main(String[] args) {
	        SwingUtilities.invokeLater(new Runnable() {
	            @Override
	            public void run() {
	                new ReadGUI();
	            }
	        });
	    }

	    public static void processEmails(String emailFilePath , String dataFilePath) throws IOException {
			
			FileInputStream emailFile = new FileInputStream(emailFilePath);
	        Workbook emailWorkbook = new XSSFWorkbook(emailFile);
	        Sheet emailSheet = emailWorkbook.getSheetAt(0);

	        FileInputStream dataFile = new FileInputStream(dataFilePath);
	        Workbook dataWorkbook = new XSSFWorkbook(dataFile);
	        Sheet dataSheet = dataWorkbook.getSheetAt(0);
	      //  Map<String, String> rowMap = new HashMap<>();
	      //  int getNumberOfNames= dataWorkbook.;
	      //  System.out.println("getNumberOfNames - "+getNumberOfNames);
	        JSONArray jsonArray = new JSONArray();
	        JSONArray jsonArray1 = new JSONArray();
	        JSONObject jsonObj = new JSONObject();
	        JSONObject jsonObj1 = new JSONObject();
	        String header = "";
	        String header1 = "";
	        Map<String, String> emailRows = new HashMap<>();
	        Row headerRow = emailSheet.getRow(0);
	       
	        for (int i = 1; i <= emailSheet.getLastRowNum(); i++) {
	            Row row = emailSheet.getRow(i);
	            if (row != null) {
	                Cell cell_data = row.getCell(5); // Assuming you are checking a specific column (index 5)
	                String cell_valuedata = cell_data == null ? "" : cell_data.getStringCellValue().trim();
	                String cellValue = "";
	                if (cell_valuedata.equalsIgnoreCase("yes")) {
	                   // JSONObject jsonObj = new JSONObject(); // Assuming you have imported org.json.JSONObject
	                    
	                    // Iterate through all cells in the row
	                    for (int j = 0; j < row.getLastCellNum(); j++) {
	                        Cell cell = row.getCell(j);
	                        header = headerRow.getCell(j).toString(); // Assuming headerRow is defined somewhere
	                       // String cellValue = "";
	                        // Check if cell is null
	                        if (cell == null) {
	                            System.out.println("Cell at row " + (i + 1) + ", column " + (j + 1) + " is null.");
	                            cellValue = null;
	                            jsonObj.put(header, cellValue);
	                           // continue; // Skip this cell
	                        }
	                        
	                        // Check if cell is blank
//	                        if (cell.getCellType() == CellType.BLANK) {
//	                            // Handle blank cell - you may set it to null or any default value
//	                            jsonObj.put(header, null); // Assuming you want to put an empty string for blanks
//	                          //  continue; // Move to next cell
//	                        }
	                        
	                        // Read cell value based on its type
	                       if(cell != null) {
	                        switch (cell.getCellType()) {
	                            case STRING:
	                                cellValue = cell.getStringCellValue();
	                                break;
	                            case NUMERIC:
	                                if (DateUtil.isCellDateFormatted(cell)) {
	                                    cellValue = cell.getDateCellValue().toString();
	                                } else {
	                                    cellValue = String.valueOf(cell.getNumericCellValue());
	                                }
	                                break;
	                            case BOOLEAN:
	                                cellValue = String.valueOf(cell.getBooleanCellValue());
	                                break;
	                            case FORMULA:
	                                cellValue = cell.getCellFormula();
	                                break;
	                            default:
	                                // Handle other types as needed
	                                cellValue = null; // Set default value if unknown type
	                        }
	                    }
	                        
	                        // Add cell value to JSON object
	                        jsonObj.put(header, cellValue);
	                    }
	                    
	                    // Add JSON object to array
	                    jsonArray.put(jsonObj);
	                }
	            }
	        }

	        System.out.println("arraylist - "+String.valueOf(jsonArray));
	        
	        
	   //   for(int m =0 ; m<=jsonArray.length(); m++) {
	        Row dataHeaderRow = dataSheet.getRow(1);
	        Map<String, Integer> headerRowMap = new HashMap<>();
	       
	        //System.out.println("bdcbwd - "+dataSheet.getLastRowNum());
	        for (int i = 2; i <= dataSheet.getLastRowNum(); i++) {
	            Row dataRow = dataSheet.getRow(i);
	            
	            for (Cell cell : dataRow) {
	            	header1 = dataHeaderRow.getCell(cell.getColumnIndex()).toString();
	            	headerRowMap.put(header1, cell.getColumnIndex());
	            	String cellValue = String.valueOf(cell);
//	            	String date = jsonArray.get(0).toString();
	            	jsonObj1.put(header1, cellValue);
	            	//System.out.println("data1  - "+jsonObj1.toString());     
	            }
	            jsonArray1.put(jsonObj1);
	        }
	        String name = new File("D:\\sample_data.xlsx").getName();
//	        String name = new File("D:\\IAR_28May_Sample(1).xslx").getName();
	        System.out.println("name - "+name);
	        Row headerData = dataSheet.getRow(0);
	        ObjectMapper objectMapper = new ObjectMapper();
	        JsonNode jsonArray12 = objectMapper.readTree(String.valueOf(jsonArray));
	        
	        for(JsonNode jsonNode : jsonArray12) {
	        String partners = 	jsonNode.get("Receiving Partners").asText().trim();
	        System.out.println("partners - "+partners);
	        JsonNode column = jsonNode.get("Filter Column");
	        String Filter_Column  = null;
	        if(column != null) {
	        	Filter_Column = jsonNode.get("Filter Column").asText().trim();
	        }
	       // String Filter_Column = jsonNode.get("Filter Column").asText().trim();
	        JsonNode filter = jsonNode.get("Filter Values");
	        String filter_value = null;
	        if(filter != null) {
	        	filter_value = jsonNode.get("Filter Values").asText().trim();
	        }
//	        
	        String[] Filter_Value_1  = null;   
	        if(filter_value != null && filter_value !="") {
	        Filter_Value_1 =filter_value.split(",");
	        }
	        
	        String receiving_partners = jsonNode.get("Receiving Partners").asText().trim();
	        String email_to = jsonNode.get("Email_To").asText().trim();
	        String email_cc = jsonNode.get("Email_CC").asText().trim();
	        String from = "Vedikabimbra@bdo.in";
	        //String from = "MitaliSinha@bdo.in";
	        String excelPath = "D:\\Excel_data\\";
	        File file1 = new File(excelPath);
	        if(!file1.exists()) {
	        	new File(excelPath).mkdirs();
	        }
	        String destinationPath = excelPath+receiving_partners+"_"+name;
	        FileOutputStream fos = new FileOutputStream(destinationPath);
	        Workbook destinationWorkbook = new XSSFWorkbook();
	        Sheet destinationSheet = destinationWorkbook.createSheet("Sheet1");
	        
	        Row header_row = dataSheet.getRow(0);
	        int header_length = header_row.getLastCellNum();
	        Row row1 = destinationSheet.createRow(0);
	        //==============================added============
	        copyRow(header_row, destinationSheet,0);
//	        for(int n = 0; n< header_length; n++) {
//	        String celll = getCellValueAsString(header_row.getCell(n));
//	        Cell cell = header_row.getCell(n);
//	        row1.createCell(n).setCellValue(celll);
//	        }
	        Row header_row_1 = dataSheet.getRow(1);
	        int header_length_1 = header_row.getLastCellNum();
	        Row row2 = destinationSheet.createRow(row1.getRowNum()+1);
	        copyRow(header_row_1, destinationSheet,1);
	        //new added//
	        
//	        Row headerRowd = destinationSheet.createRow(1); // 2nd row
//	        CellStyle headerCellStyle = destinationWorkbook.createCellStyle();
//	        headerCellStyle.setFillForegroundColor(IndexedColors.PURPLE.getIndex());
//	        headerCellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
//	        Font headerFont = workbook.createFont();
//	        headerFont.setBold(true);
//	        headerCellStyle.setFont(headerFont);
//	        
//	        Cell headerCell = headerRow.createCell(0);
//	        headerCell.setCellValue("Serial Number");
//	        headerCell.setCellStyle(headerCellStyle);
	        
//	        for(int n = 0; n< header_length_1; n++) {
//	        	String celll = getCellValueAsString(header_row_1.getCell(n));
//	        	String s_no = "S. No.";
//	        	//row2.createCell(0).setCellValue(s_no);
//	        	row2.createCell(n).setCellValue(celll);
//	        }
	        //new added//
	        
	        int rowDataCounter=2;
	      if(Filter_Column != null && filter_value != null) {
	        if(headerRowMap.containsKey(Filter_Column)) {
	        	for(int i = 2; i <= dataSheet.getLastRowNum(); i++) {
	        		
	        		Row dataRow = dataSheet.getRow(i);	
	        		
	        		Cell cel = dataRow.getCell(headerRowMap.get(Filter_Column));
	        		String cellValue = String.valueOf(cel);
	        		//if(Filter_Value.length>=0) {}
	        	for(int k =0; k<Filter_Value_1.length;k++) {
	        		String filter_value_1 = "";
	        		filter_value_1= Filter_Value_1[k].trim();
	        	if(cellValue != null) {
	        		if(cellValue.equalsIgnoreCase(filter_value_1)) {
	        			int lastrow = destinationSheet.getLastRowNum();
	        			Row dataRow_1 = dataSheet.getRow(i);
	        			//Row row = destinationSheet.createRow(lastrow+1);
	        			copyRow(dataRow_1, destinationSheet,rowDataCounter);
	        			rowDataCounter++;
//	        			for(int m = 0 ;m<=dataRow.getLastCellNum();m++) {
//	        				String celll = getCellValueAsString(dataRow.getCell(m));
//	        			//	System.out.println("cellll - "+ celll.toString());
//	        				
//	        				row.createCell(m).setCellValue(celll);
//	        				
//	        			}
	        			
	        		}
	        	}
	        	
	        }
	        
	        	}
	        }
	       
	      }
	      else {
	      	for(int i = 2; i <= dataSheet.getLastRowNum(); i++) {
	    		
	      			Row dataRow = dataSheet.getRow(i);
	    			int lastrow = destinationSheet.getLastRowNum();
	    			Row row = destinationSheet.createRow(lastrow+1);
	    			copyRow(dataRow, destinationSheet, row.getRowNum());
//	    			for(int m = 0 ;m<=dataRow.getLastCellNum();m++) {
//	    				String celll = getCellValueAsString(dataRow.getCell(m));
//	    			//	System.out.println("cellll - "+ celll.toString());
//	    				row.createCell(m).setCellValue(celll);
//	    				
//	    			}
	    	}
	    }
	        destinationWorkbook.write(fos);
	        fos.close();
	        destinationWorkbook.close();
	        System.out.println("end");
//	        sendMail(from, email_cc, email_to, destinationPath, receiving_partners);
	        }
	      emailWorkbook.close();
	      dataWorkbook.close();
}
	    //new//
//	    	  private static String getCellValueAsString(Cell cell) {
//	    	        String cellValue = "";
//	    	        if (cell != null) {
//	    	            switch (cell.getCellType()) {
//	    	                case STRING:
//	    	                    cellValue = cell.getStringCellValue();
//	    	                    break;
//	    	                case NUMERIC:
//	    	                    if (DateUtil.isCellDateFormatted(cell)) {
//	    	                        cellValue = cell.getDateCellValue().toString();
//	    	                    } else {
//	    	                        cellValue = String.valueOf(cell.getNumericCellValue());
//	    	                    }
//	    	                    break;
//	    	                case BOOLEAN:
//	    	                    cellValue = String.valueOf(cell.getBooleanCellValue());
//	    	                    break;
//	    	                case FORMULA:
//	    	                    cellValue = cell.getCellFormula();
//	    	                    break;
//	    	                default:
//	    	                    // Empty cell


//	    	            }
//	    	        }
//	    	        return cellValue;
//	    	    }
	    	  //new//
	    
	    /*
	    	public static void sendMail(String from, String cc, String to, String filePath, String name) {
//	    	  
//	          // SMTP server address
//	          String host = "smtp.office365.com"; // Replace with your SMTP server address
//	          // SMTP server port
//	          String port = "587"; // Replace with your SMTP server port (usually 587 for TLS/STARTTLS)
//	          // Sender's email credentials
//	          String username = ""; // Replace with sender's email address
//	          String password = ""; // Replace with sender's email password
//
//	          // Attachment file path
//	          String attachmentPath = filePath;
//
//	          // Create properties object to hold SMTP configuration
//	          Properties props = new Properties();
//	          props.put("mail.smtp.auth", "true");
//	          props.put("mail.smtp.starttls.enable", "true");
//	          props.put("mail.smtp.host", host);
//	          props.put("mail.smtp.port", port);
//
//	          // Create Session object
//	          Session session = Session.getInstance(props, new Authenticator() {
//	              protected PasswordAuthentication getPasswordAuthentication() {
//	                  return new PasswordAuthentication(username, password);
//	              }
//	          });
//
//	          try {
//	              // Create MimeMessage object
//	              Message message = new MimeMessage(session);
//	              // Set From: header field
//	              message.setFrom(new InternetAddress(from));
//	              String[] toAddresses = to.split(";");
//	              InternetAddress[] toInternetAddresses = new InternetAddress[toAddresses.length];
//	              for (int i = 0; i < toAddresses.length; i++) {
//	                  toInternetAddresses[i] = new InternetAddress(toAddresses[i].trim());
//	              }
//	              message.setRecipients(Message.RecipientType.TO, toInternetAddresses);
//	              if (cc != null && !cc.isEmpty()) {
//	                  String[] ccAddresses = cc.split(";");
//	                  InternetAddress[] ccInternetAddresses = new InternetAddress[ccAddresses.length];
//	                  for (int i = 0; i < ccAddresses.length; i++) {
//	                      ccInternetAddresses[i] = new InternetAddress(ccAddresses[i].trim());
//	              }
//	                  message.setRecipients(Message.RecipientType.CC, ccInternetAddresses);
//	              }
//	              LocalDate currentDate = LocalDate.now();
//	              DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");
//	              String formattedDate = currentDate.format(formatter);
//	              message.setSubject("Customer Ageing Report BDO India LLP As on Date "+formattedDate+" : "+name);
//	              BodyPart messageBodyPart = new MimeBodyPart();
//	              messageBodyPart.setText("Dear "+name+",\r\n\r\n"
//	              		+ "I hope this email finds you well.\r\n\r\n"
//	              		+ "I'm writing to provide you with an update on the aging data. Attached, you will find the latest aging report.\r\n"
//	              		+ "Please review the attached report at your earliest convenience and let me know if you have any questions or concerns.\r\n\r\n"
//	              		+ "Thank you for your continued collaboration.\r\n\r\n"
//	              		+ "Best regards,\r\n"
//	              		+ "Praveen Pandey\r\n\r\n"
//	              		+ "");
//
//	              MimeBodyPart attachmentPart = new MimeBodyPart();
//	              attachmentPart.attachFile(new File(attachmentPath));
//	              Multipart multipart = new MimeMultipart();
//	              multipart.addBodyPart(messageBodyPart); // Add email content
//	              multipart.addBodyPart(attachmentPart); // Add attachment
//	              message.setContent(multipart);
//	              Transport.send(message);
//	             // message.setFlag(Flags.Flag.DRAFT, true);
//	              
//	              
//	              // Save message
//	             // message.saveChanges();
//	              
//	             
//	              System.out.println("Email with attachment sent successfully.");
//	          } catch (MessagingException | IOException e) {
//	        	    e.printStackTrace();
//	          }
//	      

//
//		try
//	      {
//	         // Outlook application
//	         Outlook outlookApplication = new Outlook();
//	         
//	         OutlookRecipient currentUser = outlookApplication.getCurrentUser();
//	         
//	        // System.out.println(String.valueOf(currentUser.getName()));
//	         // Get the Outbox folder
//	         OutlookFolder draft = outlookApplication.getDefaultFolder(FolderType.DRAFTS);
//	         
//	         // Create a new mail in the outbox folder
//	         OutlookMail mail = (OutlookMail) draft.createItem(ItemType.MAIL);
//	        // File file = new File("D:\\Excel_data\\Amit Shah_sample_data.xlsx");
//	         LocalDate currentDate = LocalDate.now();
//	         DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");
//	         String formattedDate = currentDate.format(formatter);
//	         mail.setSubject("Customer Ageing Report BDO India LLP As on Date"+ formattedDate+" : "+ name);
//	         mail.setTo(to);
//	         mail.setCC(cc);
//	         mail.setBody("Dear "+name+",\r\n\r\n"
//	           		+ "I hope this email finds you well.\r\n\r\n"
//	           		+ "I'm writing to provide you with an update on the aging data. Attached, you will find the latest aging report.\r\n"
//	           		+ "Please review the attached report at your earliest convenience and let me know if you have any questions or concerns.\r\n\r\n"
//	           		+ "Thank you for your continued collaboration.\r\n\r\n"
//	           		+ "Best regards,\r\n"
//	           		+ currentUser.getName()
//	           		+ "");
//	         mail.getAttachments().add(new File(filePath));
//	         // Send the mail
//	         mail.save();
//	         //mail.setSaveSentMessageFolder(outbox);
//	        // mail.send();
//	         System.out.println("drafted...........");
//	         // Dispose the library
//	         outlookApplication.dispose();
//	      }
//	      
//	      catch(LibraryNotFoundException ex)
//	      {
//	         // If this error occurs, verify the file 'moyocore.dll' is present
//	         // in java.library.path
//	         System.out.println("The Java Outlook Library has not been found.");
//	         ex.printStackTrace();
//	      } catch (Exception e) {
//      e.printStackTrace();
//  }
//
		

    ComThread.InitSTA(); // Initialize the COM thread

    ActiveXComponent outlook = new ActiveXComponent("Outlook.Application");
    Dispatch mailSession = outlook.getProperty("Session").toDispatch();

    try {
        Dispatch folder = Dispatch.call(mailSession, "GetDefaultFolder", 16).toDispatch(); // 16 represents the Drafts folder

        Dispatch items = Dispatch.get(folder, "Items").toDispatch();
        int count = Dispatch.get(items, "Count").getInt();

        System.out.println("Total emails in Drafts: " + count);
        LocalDate currentDate = LocalDate.now();
	         DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");
	         String formattedDate = currentDate.format(formatter);
//	         mail.setSubject("Customer Ageing Report BDO India LLP As on Date"+ formattedDate+" : "+ name);
        // Create a new draft email
        Dispatch emailItem = Dispatch.call(items, "Add").toDispatch();
        Dispatch currentUser = Dispatch.call(mailSession, "CurrentUser").toDispatch();
        String userName = Dispatch.get(currentUser, "Name").getString();//.formatted();
//        String userName = Dispatch.get(currentUser, "Name").getString();
        System.out.println("Current User: " + userName);
        // Set properties of the email
        Dispatch.put(emailItem, "Subject", "Customer Ageing Report BDO India LLP As on Date "+ formattedDate+" : "+ name);
        Dispatch.put(emailItem, "Body",
        		"Dear "+name+",\r\n"
	           		+ "I hope this email finds you well.\r\n\r\n"
	           		+ "I'm writing to provide you with an update on the aging data. Attached, you will find the latest aging report.\r\n"
	           		+ "Please review the attached report at your earliest convenience and let me know if you have any questions or concerns.\r\n\r\n"
	           		+ "Thank you for your continued collaboration.\r\n\r\n"
	           		+ "Best regards,\r\n"
	           		+ userName
	           		+ "");

        // Fetch current user name
        
        
        String attachmentPath = filePath; // Replace with your file path
        Dispatch attachments = Dispatch.get(emailItem, "Attachments").toDispatch();
        Dispatch.call(attachments, "Add", attachmentPath);

        // Add recipients to TO field
        String toRecipients = to;
        Dispatch.put(emailItem, "To", toRecipients.toString());
       // addRecipients(emailItem, "To", toRecipients);

        // Add recipients to CC field
        String ccRecipients = cc;
        Dispatch.put(emailItem, "Cc", ccRecipients.toString());
       // addRecipients(emailItem, "Cc", ccRecipients);

        // Save the email to Drafts folder
        Dispatch.call(emailItem, "Save");

        System.out.println("Draft email with multiple recipients saved successfully.");

    } catch (Exception e) {
        e.printStackTrace();
    } finally {
        ComThread.Release(); // Release the COM thread
    }

	
	    	}
	    	
	    	
	    	*/
	    	private static void copyRow(Row sourceRow, Sheet targetSheet, int rowNum) {
	            // Create a new row in the target sheet
	    		//sourceRow.setZeroHeight(false);
	            Row newRow = targetSheet.createRow(rowNum);
	           // newRow.getZeroHeight(false);
	            
	            
	            // Iterate through cells in the source row and copy them to the new row
	            for (int col = sourceRow.getFirstCellNum(); col < sourceRow.getLastCellNum(); col++) {
	                Cell sourceCell = sourceRow.getCell(col);
	                Cell newCell = newRow.createCell(col+1);
//	                Cell newCell = newRow.createCell(col);
	              
	                if (sourceCell != null) {
	                    
	                    // Copy cell value
	                    switch (sourceCell.getCellType()) {
	                        case STRING:
	                            newCell.setCellValue(sourceCell.getStringCellValue());
	                            break;
	                        case NUMERIC:
	                            newCell.setCellValue(sourceCell.getNumericCellValue());
	                            break;
	                        case BOOLEAN:
	                            newCell.setCellValue(sourceCell.getBooleanCellValue());
	                            break;
	                        case FORMULA:
	                            newCell.setCellFormula(sourceCell.getCellFormula());
	                            break;
	                        case BLANK:
	                            // Handle blank cells if needed
	                            newCell.setCellType(CellType.BLANK);
	                            break;
	                        default:
	                            // Handle other cell types if necessary
	                            break;
	                    }

	                    // Copy cell style
	                    CellStyle sourceStyle = sourceCell.getCellStyle();
	                    CellStyle newStyle = targetSheet.getWorkbook().createCellStyle();
	                    newStyle.cloneStyleFrom(sourceStyle);
	                    newStyle.setAlignment(HorizontalAlignment.LEFT);
	                    
	                    newCell.setCellStyle(newStyle);
	                }
	            }
	        }
	

		 
	 
	
	 
		
		
	}
	 
	 

	
	

