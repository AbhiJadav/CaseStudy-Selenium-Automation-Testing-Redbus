package selenium;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.concurrent.TimeUnit;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;


public class redbus_automation {

	private static WebDriver driver;
	private static String baseUrl;
	private static String fileName="D:\\input.xls";

	public static void main(String[] args) throws Exception {


		//get data from EXCEL SHEET
		String[][] input=readExcelData(fileName,"Sheet1");

		//List of type hashmap each hashmap in this list represents input for one unique case
		List<HashMap<String, String>> list=new ArrayList<HashMap<String,String>>();

		for (int i = 0; i < (input.length-1); i++) {
			HashMap<String, String> hm=new HashMap<String, String>();
			for (int j = 0; j < input[i].length; j++) {

				switch (j) {
				case 0:
					hm.put("source", input[i][j]);
					break;

				case 1:
					hm.put("destination", input[i][j]);
					break;

				case 2:
					hm.put("departureDate", input[i][j]);
					break;

				default:
					System.err.println("Error: error while fetching the data");
					System.exit(0);
					break;
				}
			}
			list.add(hm);
		}


		//loading firefox driver
		driver = new FirefoxDriver();
		driver.manage().timeouts().implicitlyWait(4, TimeUnit.SECONDS);
		baseUrl = "https://www.redbus.in/";

		for (int i = 0; i < list.size(); i++) {
			RunRedbusAutomationTesting(
					list.get(i).get("source"),
					list.get(i).get("destination"),
					list.get(i).get("departureDate"),
					(i+1));
		}

	}



	private static void RunRedbusAutomationTesting(String from, String to, String departureDate,int testCaseNo) throws Exception {

		//declaration
		String calanderMonth,myMonth,monthName;
		int day,month,year;
		boolean current_month_calander=false;
		String[] date_in_array_format; 
		//variables needed to enter the result in excel
		List<HashMap<String, String>> outputList=new ArrayList<HashMap<String,String>>(); 

		//initialization
		try {
			date_in_array_format=departureDate.split("/");
			month=Integer.parseInt(date_in_array_format[0]);
			day=Integer.parseInt(date_in_array_format[1]);
			year=Integer.parseInt(date_in_array_format[2]);
			monthName=getMonthNameFromNum(month);
			myMonth=monthName+" "+year;
		} catch (Exception e) {
			// TODO: handle exception
			System.err.println("Error in input format of the date");
			return;
		}


		driver.get(baseUrl);
		System.out.printf("CASE "+testCaseNo+" :\n\n");
		//testing starts
		System.out.println("Automation testing Started...");
		CharSequence[] source={from};
		System.out.println("Setting source city...");
		driver.findElement(By.id("txtSource")).click();
		driver.findElement(By.id("txtSource")).clear();
		driver.findElement(By.id("txtSource")).sendKeys(source);

		//from event hash map
		HashMap<String, String> fromHM=new HashMap<String, String>();
		fromHM.put("caseNum", Integer.toString(testCaseNo));
		fromHM.put("event", "insert data into Source textBox");
		fromHM.put("input", from);
		fromHM.put("expected", "source should be set as per the input");
		fromHM.put("actual", "source is set to \""+from+"\"");
		outputList.add(fromHM);

		CharSequence[] destination={to};
		System.out.println("Setting destination city...");
		driver.findElement(By.id("txtDestination")).click();
		driver.findElement(By.id("txtDestination")).clear();
		driver.findElement(By.id("txtDestination")).sendKeys(destination);


		//To event hash map
		HashMap<String, String> toHM=new HashMap<String, String>();
		toHM.put("caseNum", Integer.toString(testCaseNo));
		toHM.put("event", "insert data into Destination textBox");
		toHM.put("input", to);
		toHM.put("expected", "destination should be set as per the input");
		toHM.put("actual", "destination is set to \""+to+"\"");
		outputList.add(toHM);

		System.out.println("Setting the date...");
		driver.findElement(By.id("txtOnwardCalendar")).click();

		do {
			calanderMonth=driver.findElement(By.cssSelector("#rbcal_txtOnwardCalendar > table.monthTable.first > tbody > tr.monthHeader > td.monthTitle")).getText();
			if(calanderMonth.equals(myMonth)){
				current_month_calander=true;
			} else {
				driver.findElement(By.cssSelector("#rbcal_txtOnwardCalendar > table.monthTable.last > tbody > tr.monthHeader > td.next > button")).click();
			}
		} while (!current_month_calander);


		String dayString=Integer.toString(day);

		if(!isToday(dayString)){
			if(!isWeDay(dayString)){
				if(!isWdDay(dayString)){
					System.err.println("Invalid input : Date is invalid");
					System.exit(0);
				}
			}
		}

		//departureDate event hash map
		HashMap<String, String> departureDateHM=new HashMap<String, String>();
		departureDateHM.put("caseNum", Integer.toString(testCaseNo));
		departureDateHM.put("event", "select date from the Calender");
		departureDateHM.put("input", departureDate);
		departureDateHM.put("expected", "date should be set according to input");
		departureDateHM.put("actual", "date is set to \""+departureDate+"\"");
		outputList.add(departureDateHM);

		driver.findElement(By.id("searchBtn")).click();
		System.out.println("Search button clicked...");
		HashMap<String, String> searchButtonHM=new HashMap<String, String>();
		searchButtonHM.put("caseNum", Integer.toString(testCaseNo));
		searchButtonHM.put("event", "click the search button");
		searchButtonHM.put("input", "");
		searchButtonHM.put("expected", "page should be redirected");
		if(driver.getCurrentUrl().equals(baseUrl)){
			searchButtonHM.put("actual", "page is not redirected");
		} else{
			searchButtonHM.put("actual", "page redirected successfully");	
		}
		outputList.add(searchButtonHM);


		StringBuffer busTypes=new StringBuffer();
		StringBuffer ratingTypes=new StringBuffer();

		System.out.println("Setting up Bus Type filters...");
		driver.findElement(By.cssSelector("#onwardSortAndFilter > div.FilterBar > div.filtersList > div.filter.BusType > a.dpBtn")).click();
		driver.findElement(By.id("BusType_AC")).click();
		busTypes.append("AC, ");
		driver.findElement(By.id("BusType_Non_AC")).click();
		busTypes.append("NonAC, ");
		driver.findElement(By.id("BusType_Sleeper")).click();
		busTypes.append("Sleeper, ");
		driver.findElement(By.id("BusType_Cab")).click();
		busTypes.append("Cab");


		//Bus Type selection event hash map
		HashMap<String, String> busTypeSelectionHM=new HashMap<String, String>();
		busTypeSelectionHM.put("caseNum", Integer.toString(testCaseNo));
		busTypeSelectionHM.put("event", "select various bus types");
		busTypeSelectionHM.put("input", "");
		busTypeSelectionHM.put("expected", " various bus types should be selected");
		busTypeSelectionHM.put("actual", "bus types \""+busTypes+"\" are selected");
		outputList.add(busTypeSelectionHM);


		System.out.println("Setting up Rating filters...");
		driver.findElement(By.cssSelector("#onwardSortAndFilter > div.FilterBar > div.filtersList > div.filter.Rating > a.dpBtn")).click();
		driver.findElement(By.id("Rating_Higher_Rated")).click();
		ratingTypes.append("High Rated, ");
		driver.findElement(By.id("Rating_All_Buses")).click();
		ratingTypes.append("All Buses ");


		//bus Rating select event hash map
		HashMap<String, String> busRatingSelectionHM=new HashMap<String, String>();
		busRatingSelectionHM.put("caseNum", Integer.toString(testCaseNo));
		busRatingSelectionHM.put("event", "select various bus ratings");
		busRatingSelectionHM.put("input", "");
		busRatingSelectionHM.put("expected", "various bus types should be selected");
		busRatingSelectionHM.put("actual", "bus types \""+ratingTypes+"\" are selected");
		outputList.add(busRatingSelectionHM);

		System.out.println("Getting least fare...");

		do {
			driver.findElement(By.linkText("Fare")).click();
		} while(!driver.findElement(By.linkText("Fare")).getAttribute("class").equals("asc"));


		String fare=driver.findElement(By.cssSelector("#wrapper1 > div.MB > div.MB > div.tripView > div.PrivateBuses > ul.BusList > li > div.busItem > div.fareBlock > span.fareSpan > span.Fare")).getText();

		//Minium fare event hash map
		HashMap<String, String> minFareHM=new HashMap<String, String>();
		minFareHM.put("caseNum", Integer.toString(testCaseNo));
		minFareHM.put("event", "fetch minium rate for \""+from+"-"+to+"\"");
		minFareHM.put("input", "");
		minFareHM.put("expected", "minimum rate should be displayed");
		minFareHM.put("actual", "minimum rate is Rs. "+fare);
		outputList.add(minFareHM);

		//print out all BUS DETAILS
		System.out.println("\nBUS DETAILS :\n");
		System.out.printf("From : "+from+"\n");
		System.out.printf("To : "+to+"\n");
		System.out.printf("Departure date : "+day+" "+monthName+","+year+"\n");
		System.out.println("Types of bus filters : "+ busTypes);
		System.out.println("Ratings of bus filters : "+ratingTypes);
		//Bus price
		System.out.println("Least fare of the bus is Rs. "+fare+"\n\n");
		
		try {
			writeExcelData(testCaseNo,"ResultSheet",outputList);	
		} catch (Exception e) {
			// TODO: handle exception
		}
		
		return;
	}

	private static boolean isWdDay(String dayString) {
		// TODO Auto-generated method stub
		List<WebElement> wdList=new ArrayList<WebElement>();
		wdList.addAll(driver.findElements(By.cssSelector("#rbcal_txtOnwardCalendar > table.monthTable.first > tbody > tr > td.wd.day")));
		for (WebElement dayElement : wdList) {
			if(dayElement.getText().equals(dayString)){
				dayElement.click();
				return true;
			}
		}
		return false;
	}


	private static boolean isWeDay(String dayString) {
		// TODO Auto-generated method stub

		List<WebElement> weList=new ArrayList<WebElement>();
		weList.addAll(driver.findElements(By.cssSelector("#rbcal_txtOnwardCalendar > table.monthTable.first > tbody > tr > td.we.day")));
		for (WebElement dayElement : weList) {
			if(dayElement.getText().equals(dayString)){
				dayElement.click();
				return true;
			}
		}

		return false;
	}


	private static boolean isToday(String dayString) {
		// TODO Auto-generated method stub
		try {
			WebElement todayElement=driver.findElement(By.cssSelector("#rbcal_txtOnwardCalendar > table.monthTable.first > tbody > tr > td.current.day"));
			if(todayElement.getText().equals(dayString)){
				todayElement.click();
				return true;
			}
		} catch (Exception e) {
			// TODO: handle exception
			return false;
		}

		return false;
	}


	private static String getMonthNameFromNum(int monthNum) {

		switch (monthNum) {

		case 1:
			return "Jan";

		case 2:
			return "Feb";

		case 3:
			return "Mar";

		case 4:
			return "Apr";

		case 5:
			return "May";
		case 6:
			return "Jun";

		case 7:
			return "July";

		case 8:
			return "Aug";

		case 9:
			return "Sept"; 

		case 10:
			return "Oct";

		case 11:
			return "Nov";

		case 12:
			return "Dec";

		default:
			return null;
		}
	}

	public static String[][] readExcelData(String fileName, String sheetName) {
		String[][] arrayExcelData = null;
		try {
			FileInputStream fs = new FileInputStream(fileName);
			Workbook wb = Workbook.getWorkbook(fs);
			Sheet sh = wb.getSheet(sheetName);

			int totalCols = sh.getColumns();
			int totalRows = sh.getRows();

			arrayExcelData = new String[totalRows][totalCols];

			for (int i= 1 ; i < totalRows; i++) {
				for (int j=0; j < totalCols; j++) {
					arrayExcelData[i-1][j] = sh.getCell(j, i).getContents();
				}
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
			e.printStackTrace();
		} catch (BiffException e) {
			e.printStackTrace();
		}
		return arrayExcelData;
	}

	private static void writeExcelData(int testCaseNo,String sheetName, List<HashMap<String, String>> outputList) {
		// TODO Auto-generated method stub
		int row=1;
		try {
			FileOutputStream os=new FileOutputStream("D:\\TestCase"+testCaseNo+".xls");
			WritableWorkbook workbook = Workbook.createWorkbook(os);
			WritableSheet sheet = workbook.createSheet(sheetName, 0);

			Label caseNumTitle = new Label(0,0, "CaseNo"); 
			sheet.addCell(caseNumTitle); 
			Label eventTitle = new Label(1,0,"What event is going to take place?"); 
			sheet.addCell(eventTitle); 
			Label inputTitle = new Label(2,0,"What should be the input"); 
			sheet.addCell(inputTitle); 
			Label expectedResultTitle = new Label(3,0," What is expected output?"); 
			sheet.addCell(expectedResultTitle); 
			Label actualResultTitle = new Label(4,0,"What is actual output"); 
			sheet.addCell(actualResultTitle);	
			
			for (int i = 0; i < outputList.size(); i++) {
				Label caseNum = new Label(0,row, outputList.get(i).get("caseNum")); 
				sheet.addCell(caseNum); 
				Label event = new Label(1,row, outputList.get(i).get("event")); 
				sheet.addCell(event); 
				Label input = new Label(2,row, outputList.get(i).get("input")); 
				sheet.addCell(input); 
				Label expectedResult = new Label(3,row, outputList.get(i).get("expected")); 
				sheet.addCell(expectedResult); 
				Label actualResult = new Label(4,row, outputList.get(i).get("actual")); 
				sheet.addCell(actualResult);	
				row++;
			}

			//Write and close the workbook
			workbook.write();
			workbook.close();

		} catch (IOException e) {
			e.printStackTrace();
		} catch (RowsExceededException e) {
			e.printStackTrace();
		} catch (WriteException e) {
			e.printStackTrace();
		}
	}

}
