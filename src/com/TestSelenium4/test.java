package com.TestSelenium4;



package codebase.TestScripts;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.Set;

import org.apache.commons.io.FileUtils;
import org.apache.log4j.Priority;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

public class Driver {
	static int rowCount=0;
public String getValue(String key,String type) throws IOException{
   	   // XSSFWorkbook objBook = new XSSFWorkbook(new FileInputStream("C:\\Users\\A631020\\selenium\\objectRepository.xlsx"));
	XSSFWorkbook objBook = new XSSFWorkbook(new FileInputStream("C:\\Users\\A631020\\workspace\\Assignment10\\src\\codebase\\Repository\\objectRepository.xlsx"));
	   XSSFSheet objSheet = objBook.getSheet("Sheet1");
	    XSSFRow objRow;
	    int RowNum = 0;
	    int ColNum = 0;
	    String strValue = null;
	    XSSFCell objCell;
	    //Iterating All the rows
	    Iterator<Row> objRowItr = objSheet.rowIterator();
	    Iterator<Cell> objCellItr;
	    while(objRowItr.hasNext()){
	            objRow = (XSSFRow) objRowItr.next();
	            RowNum++;
	            //Iterating the cells
	            objCellItr = objRow.cellIterator();
	            ColNum = 0;
	            while(objCellItr.hasNext())
	            {
	                    objCell = (XSSFCell) objCellItr.next();
	                    ColNum++;
	            }
	       }
	    int status = 0;
	    //System.out.println("Row count::"+RowNum+"  Column Count :: "+ColNum);
	    for(int i = 1;i < RowNum; i++)
	    {
	            objRow = objSheet.getRow(i);
	            //System.out.println(objRow.getCell(0).getStringCellValue());
	            if(objRow.getCell(0).getStringCellValue().equals(key))
	            {
	                    //System.out.println(objRow.getCell(0).getStringCellValue());
	                    if(objRow.getCell(1).getStringCellValue().equals(type))
	                    {
	                            strValue= objRow.getCell(2).getStringCellValue();
	                            status = 1;
	                    }
	                  }
	            if(status == 1)
	            {
	                    //System.out.println("Value Success");
	                    break;
	            }
	           }
	   return strValue;
	}
	
	
	
public String getData(String key) throws IOException{
   	 	String value = null;
	Cell cell=null ;
	try{
		
		//FileInputStream file = new FileInputStream(new File("C:\\Users\\A631020\\selenium\\TestDataFile.xlsx"));
		FileInputStream file = new FileInputStream(new File("C:\\Users\\A631020\\workspace\\Assignment10\\src\\codebase\\TestData\\TestDataFile.xlsx"));
		   
		XSSFWorkbook workbook = new XSSFWorkbook(file);
			      XSSFSheet spreadsheet = workbook.getSheetAt(0);
			     
			      XSSFRow row=null;
			      int intRowNum=0;
			      int intColNum=0;
			      
			      Iterator < Row > rowIterator = spreadsheet.iterator();
			     
			      while (rowIterator.hasNext()) 
			      {
			          row = (XSSFRow) rowIterator.next();
			         intRowNum++;
			         Iterator < Cell >  cellIterator = row.cellIterator();
			         while ( cellIterator.hasNext()) 
			         {
			            cell = cellIterator.next();
			            intColNum++;
			         }
			      }
			     
			      for(int i=0;i<intRowNum;i++){
			    	  row = spreadsheet.getRow(i);
			    	  if(row.getCell(0).getStringCellValue().equals(key)){
			    		  Iterator<Cell> cellIterator = row.cellIterator();
			                 
			                while (cellIterator.hasNext()) 
			                {
			                   cell = cellIterator.next();
			                    //Check the cell type and format accordingly
			                   if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			                	  
			                       cell.setCellType(Cell.CELL_TYPE_STRING);
			                       value = cell.getStringCellValue();
			                   }
			                   else if(cell.getCellType() == Cell.CELL_TYPE_STRING){
			                	   value = cell.getStringCellValue();
			                   }
			                }  
			    	  }
			      }
		}catch(Exception ex){
		ex.printStackTrace();
	}
	return value;
}

	
		
	public void insertInExcel(Object[] bookData,XSSFWorkbook workbook1,File copyLink){
		//System.out.println("inside insert excel");
		CreationHelper createHelper = workbook1.getCreationHelper();  
		 XSSFCellStyle hlinkstyle = workbook1.createCellStyle();
	      XSSFFont hlinkfont = workbook1.createFont();
	      hlinkfont.setUnderline(XSSFFont.U_SINGLE);
	      hlinkfont.setColor(HSSFColor.BLUE.index);
	      hlinkstyle.setFont(hlinkfont);
	      XSSFHyperlink link = (XSSFHyperlink)createHelper.createHyperlink(Hyperlink.LINK_URL);	      
	      Cell cell=null;
	    XSSFSheet spreadsheet1 = workbook1.getSheetAt(0);
         XSSFRow row = spreadsheet1.createRow(rowCount++); 
          int columnCount = 0;
          for (Object field : bookData) {
           cell = row.createCell(columnCount++);
              if (field instanceof String) {
            	  System.out.println((String) field);
                  cell.setCellValue((String) field);
              } else if (field instanceof Integer) {
                  cell.setCellValue((Integer) field);
            }
              link.setAddress(copyLink.toURI().toString());
            cell.setHyperlink((XSSFHyperlink) link);
            cell.setCellStyle(hlinkstyle);
      }
	}

	public XSSFWorkbook excelWrite() throws IOException{
		XSSFWorkbook workbook1 = new XSSFWorkbook(); 
	      //Create a blank sheet
	      XSSFSheet spreadsheet1 = workbook1.createSheet("Test Result");
	      XSSFRow row;
	      Object[] bookData ={"TestId", "TestName","Expected Output","Actual output","Status","ScreenShot"};
	                        
	      insertInExcel(bookData,workbook1,new File("C:\\Users\\A631020\\workspace\\Assignment10\\src\\codebase\\Result\\screenshot"));
		Date date = new Date() ;
		SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss") ;
		File file1 = new File("C:\\Users\\A631020\\workspace\\Assignment10\\src\\codebase\\Result\\"+dateFormat.format(date) + ".xlsx") ;
		file1.createNewFile();
		FileOutputStream file2= new FileOutputStream(file1);
		workbook1.write(file2);
	    return workbook1;
	}
	
		WebDriver driver=null; 
	
	@Parameters({"driverSelect","url"})
	@BeforeClass
   // public void beforeTest(){ 
	public void returnDriver(String driverSelect,String url) throws IOException{
		System.out.println("Run Before each test method");
		
		if(driverSelect.equals("Mozilla")) 
        { driver = new FirefoxDriver(); 
        } 
        else if(driverSelect.equals("Chrome")) 
        { 
            String path = "C:\\Users\\A631020\\selenium\\chromedriver.exe";
            String data= "webdriver.chrome.driver"; 
            System.setProperty(data,path);
            driver = new ChromeDriver(); 
        } 
        else{
        	System.setProperty("webdriver.ie.driver","C:\\Users\\A631020\\selenium\\IEDriverServer.exe");
            DesiredCapabilities capabilities = DesiredCapabilities.internetExplorer();
            capabilities.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS,true);
            driver = new InternetExplorerDriver(capabilities);
        }
	//	WebDriver driver = returnDriver(driverSelect); 
        driver.get(url); 
       
    } 
	
	
    @Test (priority=1)
public void runTest() throws InterruptedException, IOException{ 

           		System.out.println("inside test 1");
           		XSSFWorkbook workbook1=null;
        		String value[] = null;
        		Date date = new Date() ;
        		SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss") ;
        		workbook1 = excelWrite();
        		String link=null;
        		File src2=null;
        		try{
                System.out.println("on page1");
                driver.findElement(By.xpath(getValue("signIn","xpath"))).click();
               System.out.println("on page2");
               driver.findElement(By.id(getValue("emailId","id"))).sendKeys(getData("emailId"));
               driver.findElement(By.xpath(getValue("createAcc","xpath"))).click();
               
               //on register page
               Thread.sleep(10000);
              
               System.out.println("reg page");
               driver.findElement(By.xpath(getValue("mrs","xpath"))).click();
               driver.findElement(By.xpath(getValue("firstName","xpath"))).sendKeys("Sangeeta");
               driver.findElement(By.xpath(getValue("lastName","xpath"))).sendKeys("Kumari");
               //driver.findElement(By.xpath(getValue("passWord","xpath"))).sendKeys("123456789");
               System.out.println(getData("passWord"));
               driver.findElement(By.xpath(getValue("passWord","xpath"))).sendKeys(getData("passWord"));
               
               Thread.sleep(10);
               WebElement mySelectElement;
               Select dropdown;
              // WebElement mySelectElement = driver.findElement(By.xpath(getValue("day","xpath")));
              
               mySelectElement = driver.findElement(By.id("days"));
                       dropdown= new Select(mySelectElement);
               dropdown.selectByValue("8");
               
               mySelectElement = driver.findElement(By.xpath(getValue("month","xpath")));
               dropdown= new Select(mySelectElement);
               dropdown.selectByValue("4");
               
               mySelectElement = driver.findElement(By.xpath(getValue("year","xpath")));
               dropdown= new Select(mySelectElement);
               dropdown.selectByValue("1994");
               
               driver.findElement(By.xpath(getValue("companyName","xpath"))).sendKeys("Atos");
               driver.findElement(By.xpath(getValue("address","xpath"))).sendKeys("wakad");
               driver.findElement(By.xpath(getValue("city","xpath"))).sendKeys("Pune");
               mySelectElement =driver.findElement(By.xpath(getValue("state","xpath")));
               dropdown= new Select(mySelectElement);
               dropdown.selectByVisibleText("Alaska");
               
               driver.findElement(By.xpath(getValue("pinCode","xpath"))).sendKeys("45201");
               driver.findElement(By.xpath(getValue("phoneNum","xpath"))).sendKeys("963258741");
               driver.findElement(By.xpath(getValue("address1","xpath"))).clear();
               driver.findElement(By.xpath(getValue("address1","xpath"))).sendKeys("wakad1");
               driver.findElement(By.xpath(getValue("register","xpath"))).click();
               
               //register sucessfull
               if(driver.findElement(By.xpath(getValue("signOut","xpath"))).isDisplayed())
               {
            	   System.out.println("sign in verified");
            	 src2= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
           		link = "C:\\Users\\A631020\\workspace\\Assignment10\\src\\codebase\\Result\\screenshot\\s1.png";
           		FileUtils.copyFile(src2, new File(link));	
           		Object[]  bookData ={"TC01", "sign in verification","sign out link should be shown","Shown","Pass",link};
           			
           		insertInExcel(bookData,workbook1,new File(link));
               }
               driver.findElement(By.xpath(getValue("signOut","xpath"))).click();
               System.out.println("sign out sign in agn");
    } catch (IOException e) {
    	// TODO Auto-generated catch block
    	e.printStackTrace();
    }
    finally

    {        
    	dateFormat=new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");

    File file=new File("C:\\Users\\A631020\\workspace\\Assignment10\\src\\codebase\\Result\\"+dateFormat.format(date) + ".xlsx");

    FileOutputStream objfile=new FileOutputStream(file);

    workbook1.write(objfile);

    objfile.close();

        

    }
       		
              
} 

      static String textagain="abc";

    @Test (dependsOnMethods = { "runTest" },priority=2)
    public void signin() throws InterruptedException, IOException{ 
    	System.out.println("sigin");
    	driver.findElement(By.xpath(getValue("signInEmail","xpath"))).sendKeys(getData("emailId"));
        driver.findElement(By.xpath(getValue("signInPassword","xpath"))).sendKeys(getData("passWord"));
        driver.findElement(By.xpath(getValue("logIn","xpath"))).click();
        
        driver.findElement(By.xpath(getValue("dress","xpath"))).click();
        
         textagain=driver.findElement(By.xpath(getValue("dressName", "xpath"))).getText();
           System.out.println(textagain);   

        driver.findElement(By.xpath(getValue("dressSub","xpath"))).click();
       // driver.findElement(By.xpath(getValue("continueShopping","xpath"))).click();
        //driver.findElement(By.xpath(getValue("dressSub","xpath"))).click();
        
        String parentWindowHandler = driver.getWindowHandle(); // Store your parent window
        String subWindowHandler = null;

        Set<String> handles = driver.getWindowHandles(); // get all window handles
        Iterator<String> iterator = handles.iterator();
        while (iterator.hasNext()){
            subWindowHandler = iterator.next();
          }
        driver.switchTo().window(subWindowHandler);
        Thread.sleep(1000);
        //driver.findElement(By.xpath("html/body/div[1]/div[1]/header/div[3]/div/div/div[4]/div[1]/div[2]/div[4]/a/span")).click();
        driver.findElement(By.xpath(getValue("popUp","xpath"))).click();
        driver.switchTo().window(parentWindowHandler);
   
       
    }
    
    @Test (dependsOnMethods ={ "signin" },priority=3)
    public void order() throws InterruptedException, IOException{
    	
    	System.out.println("order");
    	//set quantity=2
    	XSSFWorkbook workbook1=null;
		String value[] = null;
		Date date = new Date() ;
		SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss") ;
		workbook1 = excelWrite();
		String link=null;
		File src2=null;
		try{
    	driver.findElement(By.xpath(getValue("quantity", "xpath"))).click();
    	
    	System.out.println("now verification");
  
    String dressVerf=driver.findElement(By.xpath("html/body/div[1]/div[2]/div/div[3]/div/div[2]/table/tbody/tr/td[2]/p/a")).getText();
    
    if(textagain.equals(dressVerf)){
    	System.out.println("name verified");
    }
    else {
		System.out.println("name not verified");
	}
  
   String text = driver.findElement(By.xpath("html/body/div[1]/div[2]/div/div[3]/div/div[3]/div[1]/ul/li[2]/span")).getText();
   // String text = driver.findElement(By.xpath(getData("nameVerf"))).getText();
    System.out.println(text);
    String original=((getData("fName")).concat(" ").concat(getData("lName")));
    System.out.println(original);
    
    if (text.equals(original)) {
		System.out.println("customer name verified");
		src2= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
   		link = "C:\\Users\\A631020\\workspace\\Assignment10\\src\\codebase\\Result\\screenshot\\s2.png";
   		FileUtils.copyFile(src2, new File(link));	
   		Object[]  bookData ={"TC02", "Customer Name verification","Customer Name should be shown","Shown","Pass",link};
   			
   		insertInExcel(bookData,workbook1,new File(link));
	}
    else
    {
    	System.out.println("customer name not verified");
    }
    //    text=driver.findElement(By.xpath(getValue("quantVerf", "xpath"))).getAttribute("value");
    //System.out.println(text);
  //String original=((getData("fName")).concat(" ").concat(map.get("cardMiddle")).concat(" ").concat(map.get("CardLast")).concat(map.get("Add")).concat(map.get("city")).concat(", ").concat(map.get("state")).concat(", ").concat(map.get("code")));
   // System.out.println(original);
    String textadd=driver.findElement(By.xpath("html/body/div[1]/div[2]/div/div[3]/div/div[3]/div[1]/ul/li[4]/span")).getText();		
    //System.out.println(textadd);
    if(textadd.equals(getData("address")))
    {
    	System.out.println("address verf");
    	src2= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
   		link = "C:\\Users\\A631020\\workspace\\Assignment10\\src\\codebase\\Result\\screenshot\\s3.png";
   		FileUtils.copyFile(src2, new File(link));	
   		Object[]  bookData ={"TC03", "Address verification","Address should be shown","Shown","Pass",link};
   			
   		insertInExcel(bookData,workbook1,new File(link));
	
    }
    
    //click on proceed
    driver.findElement(By.xpath(getValue("proceedToCheckout1", "xpath"))).click();
    driver.findElement(By.xpath(getValue("proceedToCheckout2", "xpath"))).click();
    driver.findElement(By.xpath(getValue("agree", "xpath"))).click();
    driver.findElement(By.xpath(getValue("proceedToCheckout3", "xpath"))).click();
    String quant=driver.findElement(By.xpath("html/body/div[1]/div[2]/div/div[3]/div/div/div[2]/table/tbody/tr/td[5]/span")).getText();
    
    if(quant.equals("2")){
    	System.out.println("quant verf");
    	src2= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
   		link = "C:\\Users\\A631020\\workspace\\Assignment10\\src\\codebase\\Result\\screenshot\\s4.png";
   		FileUtils.copyFile(src2, new File(link));	
   		Object[]  bookData ={"TC04", "quantity verification","quantity should be two","Shown","Pass",link};
   			
   		insertInExcel(bookData,workbook1,new File(link));
    	
    }
    	
    	driver.findElement(By.xpath(getValue("signout", "xpath"))).click();
    if(driver.findElement(By.xpath(getValue("signinlast", "xpath"))).isDisplayed())
    {
    	System.out.println("signout sucessfully");
    	src2= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
   		link = "C:\\Users\\A631020\\workspace\\Assignment10\\src\\codebase\\Result\\screenshot\\s5.png";
   		FileUtils.copyFile(src2, new File(link));	
   		Object[]  bookData ={"TC05", "signout verification","signin should be shown","Shown","Pass",link};
   			
   		insertInExcel(bookData,workbook1,new File(link));
	
    }
    } catch (IOException e) {
    	// TODO Auto-generated catch block
    	e.printStackTrace();
    }
    finally

    {        
    	dateFormat=new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");

    File file=new File("C:\\Users\\A631020\\workspace\\Assignment10\\src\\codebase\\Result\\"+dateFormat.format(date) + ".xlsx");

    FileOutputStream objfile=new FileOutputStream(file);

    workbook1.write(objfile);

    objfile.close();

        

    }
       		
    }

@AfterClass
public void afterTest(){ 
  System.out.println("Run After each test method"); 
 // driver.close(); 
} 

}

