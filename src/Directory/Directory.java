package Directory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;

import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.DBObject;
import com.mongodb.MongoClient;
import com.mongodb.MongoClientURI;


public class Directory {
	public static WebDriver driver;
	public static XSSFSheet sheet;
	public static String getCellVal(XSSFSheet SheetName,int i,int j) {
		XSSFCell row = SheetName.getRow(i).getCell(j);
        return row.toString();	
	}
	public static String getCellVal2(XSSFSheet SheetName,int i,int j) {
		XSSFCell row = SheetName.getRow(i).getCell(j);
		try {
			 return row.toString();
		}
		catch(Exception NullPointerException)
		{
			return null;
		
		}	
	}
	public static void CreateInstance()
	{
		System.setProperty("webdriver.chrome.driver", "C:\\Webdriver\\chromedriver.exe");
		ChromeOptions options = new ChromeOptions();
		options.addArguments("--remote-allow-origins=*");
		driver = new ChromeDriver(options);	;
		 driver.manage().deleteAllCookies();
			driver.manage().window().maximize();
			
	}
	public JSONObject Json() throws IOException, ParseException
	{
JSONParser jsonparser=new JSONParser();
	FileReader reader=new FileReader("C:\\Users\\vilas\\eclipse-workspace\\HRDirectory\\Directory.json");
	Object obj=jsonparser.parse(reader);
	JSONObject files=(JSONObject)obj;
	return files;
	}
public void Login(JSONObject json) throws IOException, ParseException, InterruptedException {
		
		JSONObject file=(JSONObject)json.get("files");
		JSONObject logindata =(JSONObject) json.get("logindata");
		String chromedriver = (String) file.get("chromedriver");
		String username = (String) logindata.get("username");
		String password = (String) logindata.get("password");
		String Login = (String) logindata.get("logclick");
		String email = (String) logindata.get("Email");
		String passwordvalue = (String) logindata.get("PasswordValue");
		driver.get("https://addithr.azurewebsites.net/");
		driver.manage().window().maximize();
		Thread.sleep(2000);
		 driver.findElement(By.xpath(username)).sendKeys(email);
		 driver.findElement(By.xpath(password)).sendKeys(passwordvalue);
		 driver.findElement(By.xpath(Login)).click(); 
		 Thread.sleep(20000);
	}
public static  XSSFSheet Data_Provider() throws IOException 
{
	File file =    new File("C:\\Users\\vilas\\eclipse-workspace\\HRDirectory\\HRDirectory.xlsx");
	FileInputStream inputStream = new FileInputStream(file);
		XSSFWorkbook wb=new XSSFWorkbook(inputStream);
		  sheet=wb.getSheet("Sheet1");
		 XSSFRow roww = sheet.getRow(0);
		 int cellcount= roww.getPhysicalNumberOfCells();
		 int rowCount=sheet.getLastRowNum()+1;
			System.out.println(rowCount);
			System.out.println(cellcount);
			return sheet;
			//return new Pair(sheet, rowCount);
			
}
public static void datePicker(JSONObject json,XSSFSheet sheetname,int i,int j) throws InterruptedException {JSONObject companyprofile=(JSONObject)json.get("CompanyProfile");
JSONObject directory=(JSONObject)json.get("Directory");
String year = (String)directory .get("Year");
String next = (String)directory .get("Next");
String previous = (String)directory .get("Previous");
String month = (String)directory .get("Month");
String monthselection = (String)directory .get("MonthSelection");
String date = (String)directory .get("Date");
String ok = (String)directory .get("OK");
 Actions actions2 = new Actions(driver);
 Actions actions1 = new Actions(driver);
	actions1.moveToElement( driver.findElement(By.xpath(year))).perform();
	Thread.sleep(1000);	
	String ad= getCellVal(sheetname,i,j);
	String[] ab= ad.split("/");
	String de = ab[0];
	String nj = ab[1];
	String lm = ab[2];
String monthname= DateConverter(ad);
	int t=Integer.parseInt(lm);
String cd=	driver.findElement(By.xpath(year)).getText();
int m=Integer.parseInt(cd);
if(t>m)
{
	while(!driver.findElement(By.xpath(year)).getText().contains(lm))
		{
   driver.findElement(By.xpath(next)).click();
		}
}
else
{
	while(!driver.findElement(By.xpath(year)).getText().contains(lm))
	{
driver.findElement(By.xpath(previous)).click();
	}
}
driver.findElement(By.xpath(month)).click();
List<WebElement> Month = driver.findElements(By.xpath(monthselection));
for(int mo=0;mo<Month.size();mo++)
{	 
String mn= Month.get(mo).getText();
if(mn.equals(monthname))
{
  Month.get(mo).click();
}
}
List<WebElement> Date= driver.findElements(By.xpath(date));
for(int da=0;da<Date.size();da++)
{
String dy=  Date.get(da).getText();
if(dy.equals(de))
{
 Date.get(da).click();
break;
 }
}	
driver.findElement(By.xpath(ok)).click();
}
public static String DateConverter(String abc)
{
	String[] ab = abc.split("/");
String ce= ab[0];
String ef = ab[1];
String gh = ab[2];
String month= null;
if(ef.equals("01"))
{
	month= "January";
}
if(ef.equals("02"))
{
	month= "February";
}
if(ef.equals("03"))
{
	month= "March";
}
if(ef.equals("04"))
{
	month= "April";
}
if(ef.equals("05"))
{
	month= "May";
}
if(ef.equals("06"))
{
	month= "June";
}
if(ef.equals("07"))
{
	month= "July";
}
if(ef.equals("08"))
{
	month= "August";
}
if(ef.equals("09"))
{
	month= "Seprember";
}
if(ef.equals("10"))
{
	month= "October";
}
if(ef.equals("11"))
{
	month= "November";
}

if(ef.equals("12"))
{
	month= "December";
}
	return month;	
}
public static List<String> getdata()
{
	MongoClientURI uri = new MongoClientURI("mongodb+srv://additlabs:Addit2021@familydb.wzxxe.mongodb.net/test?authSource=admin&replicaSet=atlas-b3hasj-shard-0&readPreference=primary&ssl=true&connecTimeoutMS=30000&socketTimeoutMS=30000");
	MongoClient mongoclient = new MongoClient(uri);
	DB db= mongoclient.getDB("vilasjain9611-gmail-com63819807653676");
	DBCollection coll = db.getCollection("companyId");
	DBCursor iterDoc = coll.find();
    Iterator it = iterDoc.iterator();
	DBCursor cursor = coll.find();
	List<DBObject> list = cursor.toArray();
	List<String> names = list.stream().map(o -> String.valueOf(o.get("entityType"))).collect(Collectors.toList());
		List<String> names1 = list.stream().map(o -> String.valueOf(o.get("dateOfIncorporation"))).collect(Collectors.toList());
List<String> Elements = new ArrayList<>();
		
		Elements.addAll(names);
		Elements.addAll(names1);
		return Elements;
    
    }
public static void Directory(JSONObject json,XSSFSheet sheetname) throws InterruptedException, FileNotFoundException
{
	JSONObject directory=(JSONObject)json.get("Directory");
	String directoryname = (String)directory .get("Directory");
	String addemployee = (String)directory .get("AddEmployee");
	String firstname = (String)directory .get("FirstName");
	String middlename = (String)directory .get("MiddleName");
	String lastname = (String)directory .get("LastName");
	String employeeid = (String)directory .get("EmployeeID");
	String anualctc = (String)directory .get("AnualCTC");
	String rmselect = (String)directory .get("ReportingManager");
	String roleselect = (String)directory .get("RoleSelect");
	String phonenumber = (String)directory .get("PhoneNumber");
	String emailid = (String)directory .get("EmailID");
	String password = (String)directory .get("Password");
	String next = (String)directory .get("NextButton");
	String dob = (String)directory .get("DOB");
	String gender = (String)directory .get("Gender");
	String offemail = (String)directory .get("OffEmail");
	String department = (String)directory .get("Department");
	String subdepartment = (String)directory .get("SubDepartment");
	String designation = (String)directory .get("Designation");
	String jobtitle = (String)directory .get("JobTitle");
	String worklocation = (String)directory .get("WorkLocation");
	String employeetype = (String)directory .get("EmployeeType");
	String probationperiod = (String)directory .get("ProbationPeriod");
	String dateofjoining = (String)directory .get("DateOfJoining");
	String accountholdername = (String)directory .get("AccountHolderName");
	String bankname = (String)directory .get("BankName");
	String city = (String)directory .get("City");
	String branchname = (String)directory .get("BranchName");
	String ifsccode = (String)directory .get("IFSCCode");
	String accountnumber = (String)directory .get("AccountNumber");
	String update = (String)directory .get("Update");
	FileOutputStream outputstream = new FileOutputStream("C:\\Users\\vilas\\eclipse-workspace\\HRDirectory\\HRDirectoryOutput.csv",true);
	PrintWriter pw = new PrintWriter(outputstream);
	pw.println("TEST CASE NO,TEST CASE DESCRIPITION,FIRST NAME,MIDDLE NAME,LASE NAME,EMPLOYEE ID,ANUAL CTC,REPORTING MANAGER,ROLE,PHONE NUMBER,EMAIL ID,PASSWORD,DOB,GENDER,OFF EMAIL,DEPARTMENT,SUBDEPARTMENT,DESIGNATION,JOB TITLE,WORK LOCATION,EMPLOYEE TYPE,PP,DOJ,ACC HOLDER NAME,BANK NAME,CITY,BRANCH NAME,IFSC CODE,ACC NUMBER,EXPECTED RESULT,ACTUAL RESULT,RESULT,REMARKS");
	driver.findElement(By.xpath(directoryname)).click();
	for(int i=1;i<16;i++) {
		Thread.sleep(10000);
		driver.findElement(By.xpath(addemployee)).click();
		Thread.sleep(5000);
		pw.println();
		XSSFRow row=sheet.getRow(i);
	driver.findElement(By.xpath(firstname)).sendKeys(getCellVal(sheet,i,2));
	driver.findElement(By.xpath(middlename)).sendKeys(getCellVal(sheet,i,3));
	driver.findElement(By.xpath(lastname)).sendKeys(getCellVal(sheet,i,4));
	driver.findElement(By.xpath(employeeid)).sendKeys(getCellVal(sheet,i,5));
	driver.findElement(By.xpath(anualctc)).sendKeys(getCellVal(sheet,i,6));
	String rm = getCellVal2(sheet,i,7);
	String value = getCellVal2(sheet,i,8);
	if(rm==null)
	{
		if(value==null)
		{
			driver.findElement(By.xpath(phonenumber)).sendKeys(getCellVal(sheet,i,9));
			driver.findElement(By.xpath(emailid)).sendKeys(getCellVal(sheet,i,10));
			driver.findElement(By.xpath(password)).sendKeys(getCellVal(sheet,i,11));
			Thread.sleep(1000);
		}
		if(value!=null)
		{
			Thread.sleep(1000);
			driver.findElement(By.xpath(roleselect)).click();
			Thread.sleep(1000);
			driver.findElement(By.xpath("//span[contains(text(),'"+value+"')]")).click();
			driver.findElement(By.xpath(phonenumber)).sendKeys(getCellVal(sheet,i,9));
			driver.findElement(By.xpath(emailid)).sendKeys(getCellVal(sheet,i,10));
			driver.findElement(By.xpath(password)).sendKeys(getCellVal(sheet,i,11));
			Thread.sleep(1000);
		}
	}
	if(rm!=null)
	{
		if(value==null)
		{ 
			Thread.sleep(1000);
			driver.findElement(By.xpath(rmselect)).click();
			Thread.sleep(1000);
			driver.findElement(By.xpath("//span[contains(text(),'"+rm+"')]")).click();
			driver.findElement(By.xpath(phonenumber)).sendKeys(getCellVal(sheet,i,9));
			driver.findElement(By.xpath(emailid)).sendKeys(getCellVal(sheet,i,10));
			driver.findElement(By.xpath(password)).sendKeys(getCellVal(sheet,i,11));
			Thread.sleep(1000);
		}
		if(value!=null)
		{
			Thread.sleep(1000);
		driver.findElement(By.xpath(rmselect)).click();
		Thread.sleep(1000);
		driver.findElement(By.xpath("//span[contains(text(),'"+rm+"')]")).click();
		driver.findElement(By.xpath(roleselect)).click();
		Thread.sleep(1000);
		driver.findElement(By.xpath("//span[contains(text(),'"+value+"')]")).click();
		driver.findElement(By.xpath(phonenumber)).sendKeys(getCellVal(sheet,i,9));
		driver.findElement(By.xpath(emailid)).sendKeys(getCellVal(sheet,i,10));
		driver.findElement(By.xpath(password)).sendKeys(getCellVal(sheet,i,11));
		Thread.sleep(1000);
		}
	}
	driver.findElement(By.xpath(next)).click();
	Thread.sleep(3000);
	try
	{
	try
	{
			String Error = getCellVal(sheet,i,29);
			String ActualError = driver.findElement(By.xpath("//small[contains(text(),'This field is required')]")).getText();
			if(Error.equals(ActualError))
			{
				String Result ="PASS";
				String Remarks = "Some fields are missing or blank";
				pw.println(row.getCell(0)+","+row.getCell(1)+","+row.getCell(2)+","+row.getCell(3)+","+row.getCell(4)+","+row.getCell(5)+","+row.getCell(6)+","+row.getCell(7)+","+row.getCell(8)+","+row.getCell(9)+","+row.getCell(10)+","+row.getCell(11)+","+row.getCell(12)+","+row.getCell(13)+","+row.getCell(14)+","+row.getCell(15)+","+row.getCell(16)+","+row.getCell(17)+","+row.getCell(18)+","+row.getCell(19)+","+row.getCell(20)+","+row.getCell(21)+","+row.getCell(22)+","+row.getCell(23)+","+row.getCell(24)+","+row.getCell(25)+","+row.getCell(26)+","+row.getCell(27)+","+row.getCell(28)+","+Error+","+ActualError+","+Result+","+Remarks);
			}
		}
		catch(Exception e)
		{
			try {
				String Error = getCellVal(sheet,i,29);
				String ErrorMsg = driver.findElement(By.xpath("//small[contains(text(),'Enter a valid email id')]")).getText();
				if(Error.equals(ErrorMsg))
				{
					String Result ="PASS";
					String Remarks = "Invalid Email ID";
					pw.println(row.getCell(0)+","+row.getCell(1)+","+row.getCell(2)+","+row.getCell(3)+","+row.getCell(4)+","+row.getCell(5)+","+row.getCell(6)+","+row.getCell(7)+","+row.getCell(8)+","+row.getCell(9)+","+row.getCell(10)+","+row.getCell(11)+","+row.getCell(12)+","+row.getCell(13)+","+row.getCell(14)+","+row.getCell(15)+","+row.getCell(16)+","+row.getCell(17)+","+row.getCell(18)+","+row.getCell(19)+","+row.getCell(20)+","+row.getCell(21)+","+row.getCell(22)+","+row.getCell(23)+","+row.getCell(24)+","+row.getCell(25)+","+row.getCell(26)+","+row.getCell(27)+","+row.getCell(28)+","+Error+","+ErrorMsg+","+Result+","+Remarks);
			}
	     }
			catch(Exception e1)
			{
				try {
		    	 String Error1 = getCellVal(sheet,i,29);
					String ActualError1 = driver.findElement(By.xpath("//small[contains(text(),'Password must be at least 8 characters')]")).getText();
					if(Error1.equals(ActualError1))
					{
						String Result ="PASS";
						String Remarks = "Invalid Password";
						pw.println(row.getCell(0)+","+row.getCell(1)+","+row.getCell(2)+","+row.getCell(3)+","+row.getCell(4)+","+row.getCell(5)+","+row.getCell(6)+","+row.getCell(7)+","+row.getCell(8)+","+row.getCell(9)+","+row.getCell(10)+","+row.getCell(11)+","+row.getCell(12)+","+row.getCell(13)+","+row.getCell(14)+","+row.getCell(15)+","+row.getCell(16)+","+row.getCell(17)+","+row.getCell(18)+","+row.getCell(19)+","+row.getCell(20)+","+row.getCell(21)+","+row.getCell(22)+","+row.getCell(23)+","+row.getCell(24)+","+row.getCell(25)+","+row.getCell(26)+","+row.getCell(27)+","+row.getCell(28)+","+Error1+","+ActualError1+","+Result+","+Remarks);
				}  
				}
				catch(Exception e2)
				{
					 String Error1 = getCellVal(sheet,i,29);
						String ActualError1 = driver.findElement(By.xpath("//small[contains(text(),'Please provide valid phone number')]")).getText();
						if(Error1.equals(ActualError1))
						{
							String Result ="PASS";
							String Remarks = "Invalid Phone Number";
							pw.println(row.getCell(0)+","+row.getCell(1)+","+row.getCell(2)+","+row.getCell(3)+","+row.getCell(4)+","+row.getCell(5)+","+row.getCell(6)+","+row.getCell(7)+","+row.getCell(8)+","+row.getCell(9)+","+row.getCell(10)+","+row.getCell(11)+","+row.getCell(12)+","+row.getCell(13)+","+row.getCell(14)+","+row.getCell(15)+","+row.getCell(16)+","+row.getCell(17)+","+row.getCell(18)+","+row.getCell(19)+","+row.getCell(20)+","+row.getCell(21)+","+row.getCell(22)+","+row.getCell(23)+","+row.getCell(24)+","+row.getCell(25)+","+row.getCell(26)+","+row.getCell(27)+","+row.getCell(28)+","+Error1+","+ActualError1+","+Result+","+Remarks);
					}  
				}
			}
	}
	Thread.sleep(2000);
	}
	catch(Exception e2)
	{
		String value1 = getCellVal2(sheet,i,12);
		Thread.sleep(5000);
		if(value1==null)
		{
			continue;
		}
		if(value1!=null)
		{
			driver.findElement(By.xpath(dob)).click();
			   datePicker(json,sheet,i,12);
		}
		String value2 = getCellVal2(sheet,i,13);
		Thread.sleep(5000);
		if(value2==null)
		{
			continue;
		}
		if(value2!=null)
		{
			driver.findElement(By.xpath(gender)).click();
			   String genderselection = getCellVal(sheet,i,13);
			   driver.findElement(By.xpath("//span[contains(text(),'"+genderselection+"')]")).click();
		}
   Thread.sleep(1000);
   driver.findElement(By.xpath(offemail)).sendKeys(getCellVal(sheet,i,14));
   String value3 = getCellVal2(sheet,i,15);
	if(value3==null)
	{
		continue;
	}
	if(value3!=null)
	{
		driver.findElement(By.xpath(department)).click();
		   String departmentselection = getCellVal(sheet,i,15);
		   driver.findElement(By.xpath("//span[contains(text(),'"+departmentselection+"')]")).click();
	}
	 String value4 = getCellVal2(sheet,i,16);
		if(value4==null)
		{
			continue;
		}
		if(value4!=null)
		{
			 driver.findElement(By.xpath(subdepartment)).click();
			   String subdepartmentselection = getCellVal(sheet,i,16);
			   driver.findElement(By.xpath("//span[contains(text(),'"+subdepartmentselection+"')]")).click();
		}
   JavascriptExecutor jse1 = (JavascriptExecutor)driver;
     jse1.executeScript("window.scrollBy(0,800)");
     Thread.sleep(1000);
     String value5 = getCellVal2(sheet,i,17);
 	if(value5==null)
	{
		continue;
	}
	if(value5!=null)
	{
		  driver.findElement(By.xpath(designation)).click();
		   String designationselection = getCellVal(sheet,i,17);
		   driver.findElement(By.xpath("//span[contains(text(),'"+designationselection+"')]")).click();
	}
   driver.findElement(By.xpath(jobtitle)).sendKeys(getCellVal(sheet,i,18));
   String value6 = getCellVal2(sheet,i,19);
	if(value6==null)
	{
		continue;
	}
	if(value6!=null)
	{
		driver.findElement(By.xpath(worklocation)).click();
		   String worklocationselection = getCellVal(sheet,i,19);
		   driver.findElement(By.xpath("//span[contains(text(),'"+worklocationselection+"')]")).click();
	}
	String value7 = getCellVal2(sheet,i,20);
	if(value7==null)
	{
		continue;
	}
	if(value7!=null)
	{
		 driver.findElement(By.xpath(employeetype)).click();
		   String employeetypeselection = getCellVal(sheet,i,20);
		   driver.findElement(By.xpath("//span[contains(text(),'"+employeetypeselection+"')]")).click();
	}
	String value8 = getCellVal2(sheet,i,21);
	if(value8==null)
	{
		continue;
	}
	if(value8!=null)
	{
		 driver.findElement(By.xpath(probationperiod)).click();
		   String probationperiodselection = getCellVal(sheetname,i,21);
		   driver.findElement(By.xpath("//span[contains(text(),'"+probationperiodselection+"')]")).click();
	}
	String value9 = getCellVal2(sheet,i,22);
	if(value9==null)
	{
		continue;
	}
	if(value9!=null)
	{
		 driver.findElement(By.xpath(dateofjoining)).click();
		   datePicker(json,sheet,i,22);
	}
   driver.findElement(By.xpath(accountholdername)).sendKeys(getCellVal(sheet,i,23));
   driver.findElement(By.xpath(bankname)).sendKeys(getCellVal(sheet,i,24));
   driver.findElement(By.xpath(city)).sendKeys(getCellVal(sheet,i,25));
   driver.findElement(By.xpath(branchname)).sendKeys(getCellVal(sheet,i,26));
   driver.findElement(By.xpath(ifsccode)).sendKeys(getCellVal(sheet,i,27));
   driver.findElement(By.xpath(accountnumber)).sendKeys(getCellVal(sheet,i,28));
   driver.findElement(By.xpath(update)).click();
   Thread.sleep(4000);
   try {
		String Error = getCellVal(sheet,i,29);
		String ActualError = driver.findElement(By.xpath("//div[contains(text(),'User with provided email id is already exits')]")).getText();
		if(Error.equals(ActualError))
		{
			String Result ="PASS";
			String Remarks = "User Already Exists";
			pw.println(row.getCell(0)+","+row.getCell(1)+","+row.getCell(2)+","+row.getCell(3)+","+row.getCell(4)+","+row.getCell(5)+","+row.getCell(6)+","+row.getCell(7)+","+row.getCell(8)+","+row.getCell(9)+","+row.getCell(10)+","+row.getCell(11)+","+row.getCell(12)+","+row.getCell(13)+","+row.getCell(14)+","+row.getCell(15)+","+row.getCell(16)+","+row.getCell(17)+","+row.getCell(18)+","+row.getCell(19)+","+row.getCell(20)+","+row.getCell(21)+","+row.getCell(22)+","+row.getCell(23)+","+row.getCell(24)+","+row.getCell(25)+","+row.getCell(26)+","+row.getCell(27)+","+row.getCell(28)+","+Error+","+ActualError+","+Result+","+Remarks);
		}
	}
	catch(Exception e)
	{
		try {
			String Success = getCellVal(sheet,i,29);
		String SuccessMsg = driver.findElement(By.xpath("//div[contains(text(),'Employee onboarded successfully,')]")).getText();
		if(Success.equals(SuccessMsg))
		{
			String Result ="PASS";
			String Remarks = "New Employee Onboarded successfully";
			String Success1= "Employee onboarded successfully:";
			pw.println(row.getCell(0)+","+row.getCell(1)+","+row.getCell(2)+","+row.getCell(3)+","+row.getCell(4)+","+row.getCell(5)+","+row.getCell(6)+","+row.getCell(7)+","+row.getCell(8)+","+row.getCell(9)+","+row.getCell(10)+","+row.getCell(11)+","+row.getCell(12)+","+row.getCell(13)+","+row.getCell(14)+","+row.getCell(15)+","+row.getCell(16)+","+row.getCell(17)+","+row.getCell(18)+","+row.getCell(19)+","+row.getCell(20)+","+row.getCell(21)+","+row.getCell(22)+","+row.getCell(23)+","+row.getCell(24)+","+row.getCell(25)+","+row.getCell(26)+","+row.getCell(27)+","+row.getCell(28)+","+Success1+","+Success1+","+Result+","+Remarks);
	}
			}
		catch(Exception e1)
		{
			 try {
		    		 String Error = getCellVal(sheet,i,29);
					String ErrorMsg = driver.findElement(By.xpath("//div[contains(text(),'Employee age should be greater or equal to 18 years')]")).getText();
					if(Error.equals(ErrorMsg))
					{
						String Result ="PASS";
						String Remarks = "Employee Age is less than 18 Years";
						pw.println(row.getCell(0)+","+row.getCell(1)+","+row.getCell(2)+","+row.getCell(3)+","+row.getCell(4)+","+row.getCell(5)+","+row.getCell(6)+","+row.getCell(7)+","+row.getCell(8)+","+row.getCell(9)+","+row.getCell(10)+","+row.getCell(11)+","+row.getCell(12)+","+row.getCell(13)+","+row.getCell(14)+","+row.getCell(15)+","+row.getCell(16)+","+row.getCell(17)+","+row.getCell(18)+","+row.getCell(19)+","+row.getCell(20)+","+row.getCell(21)+","+row.getCell(22)+","+row.getCell(23)+","+row.getCell(24)+","+row.getCell(25)+","+row.getCell(26)+","+row.getCell(27)+","+row.getCell(28)+","+Error+","+ErrorMsg+","+Result+","+Remarks);
				}
	    		 }
			 catch(Exception e3)
			 {
				 try
				 {
	    		 String Error = getCellVal(sheet,i,29);
				String ErrorMsg = driver.findElement(By.xpath("//div[contains(text(),'Employee onboarding failed, Employee basic info update failed, Employee ID might already exists')]")).getText();
				if(Error.equals(ErrorMsg))
				{
					String Error1 = "Employee onboarding failed: Employee basic info update failed: Employee ID might already exists";
					String Result ="PASS";
					String Remarks = "Employee ID is already Exists";
					pw.println(row.getCell(0)+","+row.getCell(1)+","+row.getCell(2)+","+row.getCell(3)+","+row.getCell(4)+","+row.getCell(5)+","+row.getCell(6)+","+row.getCell(7)+","+row.getCell(8)+","+row.getCell(9)+","+row.getCell(10)+","+row.getCell(11)+","+row.getCell(12)+","+row.getCell(13)+","+row.getCell(14)+","+row.getCell(15)+","+row.getCell(16)+","+row.getCell(17)+","+row.getCell(18)+","+row.getCell(19)+","+row.getCell(20)+","+row.getCell(21)+","+row.getCell(22)+","+row.getCell(23)+","+row.getCell(24)+","+row.getCell(25)+","+row.getCell(26)+","+row.getCell(27)+","+row.getCell(28)+","+Error1+","+Error1+","+Result+","+Remarks);
			}
    		 }
					catch(Exception e4)
    		     {
	    	 String Result ="FAIL";
				String Remarks = "Something went wrong/Error Message has been mismatched";
				pw.println(row.getCell(0)+","+row.getCell(1)+","+row.getCell(2)+","+row.getCell(3)+","+row.getCell(4)+","+row.getCell(5)+","+row.getCell(6)+","+row.getCell(7)+","+row.getCell(8)+","+row.getCell(9)+","+row.getCell(10)+","+row.getCell(11)+","+row.getCell(12)+","+row.getCell(13)+","+row.getCell(14)+","+row.getCell(15)+","+row.getCell(16)+","+row.getCell(17)+","+row.getCell(18)+","+row.getCell(19)+","+row.getCell(20)+","+row.getCell(21)+","+row.getCell(22)+","+row.getCell(23)+","+row.getCell(24)+","+row.getCell(25)+","+row.getCell(26)+","+row.getCell(27)+","+row.getCell(28)+","+Result+","+Remarks);
    		     }
    		     }
		}
		}
}
	Thread.sleep(2000);
	driver.findElement(By.xpath(directoryname)).click();
	}
	pw.close();		 
	System.out.println("Completed");
	System.out.println("2nd Completed");
        System.out.println("3rd Completed");
        System.out.println("4th Completed");
}
	public static void main(String[] args) throws InterruptedException, IOException, ParseException {
		// TODO Auto-generated method stub
		Directory obj = new Directory();
		obj.CreateInstance();
		sheet=  obj.Data_Provider();
		JSONObject json=  obj.Json();
		obj.Login(json);
		obj.Directory(json, sheet);

	}

}
//User with provided email id is already exits
//Employee onboarded successfully,
//Employee age should be greater or equal to 18 years