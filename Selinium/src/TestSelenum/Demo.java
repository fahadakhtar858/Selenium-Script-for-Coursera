package TestSelenum;

import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import java.io.*;
import java.nio.file.Path;

public class Demo {

	@SuppressWarnings("resource")
	public static void main(String[] args) {
		
		//WebDriver driver = new ChromeDriver();
		
		// TODO Auto-generated method stub
		String topic ="";//"BlockChain";
		String type = "";//"course";
		String path = "/Users/Fahad/Desktop/Test01/";//"//Users//Fahad//Desktop//";
		//courseraCrawler(topic);
		
		System.out.print("Enter the Topic: ");
		Scanner sc = new Scanner(System.in);
		topic = sc.nextLine();
		System.out.print("Enter the Path to save the Excel file: ");
		path = sc.nextLine();
		
		int pageNum = 0, courseNum = 0, pages = 0;
		
		System.out.print("Enter the page number of the search results, the default value is zero:");
		pageNum = sc.nextInt();
		
		System.out.print("Enter the course number from the search result page, the default value is zero:");
		courseNum = sc.nextInt();
		
		System.out.print("Enter the number of pages you want to Crawl:");
		pages = sc.nextInt();
		
		
		
		//courseraCourses(topic, path, pageNum, courseNum, pages);
		edxCourses(topic, path, pageNum, courseNum, pages);
		
		
	}
	
	public static void edxCourseFile(ArrayList<Course> courses, String path,String topic) {
		try {
			String FileName = path+"EDx-"+topic+".xlsx";
			File file = new File(FileName);
			HSSFWorkbook wb;
			System.out.print("Createing the Excel File.");

			
			if(file.exists()) {
				FileInputStream inputStream = new FileInputStream(file);
				wb = new HSSFWorkbook(inputStream);
				//System.out.print("File is writeable\n");
				HSSFSheet sheet = wb.getSheet(topic);
				int rowcount = sheet.getLastRowNum();
				
				HSSFRow row;
		    	int i = rowcount;
		    	
		    	for(Course obj : courses) {
		    		row = sheet.createRow(++i);
		    		String Instructors ="";
		    		//String Departments ="";
		    		String Designation ="";
		    		ArrayList<Instructor>list = obj.instructors;
		    		for(int a =0;a<list.size();a++) {
		    			Instructor one = list.get(a);
		    			Instructors = Instructors+one.instructorName+", ";
		    			//Departments = one.instructorDepartment;
		    			Designation = Designation+one.instructorDesigination+", ";
		    		}
		    		//System.out.println(Instructors+"\n");
		    		//System.out.println(Departments+"\n");
		    		//System.out.println(Designation+"\n");
		    		
		    		row.createCell(0).setCellValue(i);  
			    	row.createCell(1).setCellValue(obj.courseName); 
			    	row.createCell(2).setCellValue(obj.url);
			    	row.createCell(3).setCellValue(Instructors);
			    	row.createCell(4).setCellValue(obj.offeredBy);
			    	row.createCell(5).setCellValue(Designation);
			    	//row.createCell(6).setCellValue(Departments);
			    	row.createCell(6).setCellValue(obj.skillsoffered);  
			    	//row.createCell(8).setCellValue(obj.ratings);
			    	//row.createCell(9).setCellValue(obj.reviews);
			    	row.createCell(7).setCellValue(obj.noOfEnrollments);
			    	row.createCell(8).setCellValue(obj.level);
			    	row.createCell(9).setCellValue(obj.courseType);
			    	row.createCell(10).setCellValue(obj.duration);
		    		
		    		
		    	}
		    	inputStream.close();
				
				FileOutputStream fileOut = new FileOutputStream(FileName);
				wb.write(fileOut);
				fileOut.close();
				wb.close();
			}
			else {
				File file01 = new File(path);
				file01.mkdir();
				
				wb = new HSSFWorkbook();
		    	String Filename = path+"EDx-"+topic+".xlsx";
		    	
		    	HSSFSheet sheet = wb.createSheet(topic);
		    	
		    	HSSFRow rowhead = sheet.createRow(0);
		    	
		    	rowhead.createCell(0).setCellValue("S.No.");  
		    	rowhead.createCell(1).setCellValue("Course Name"); 
		    	rowhead.createCell(2).setCellValue("URL");
		    	rowhead.createCell(3).setCellValue("Instructor");
		    	rowhead.createCell(4).setCellValue("Institute");
		    	rowhead.createCell(5).setCellValue("Designation");
		    	//rowhead.createCell(6).setCellValue("Department");
		    	rowhead.createCell(6).setCellValue("Skills");  
		    	//rowhead.createCell(8).setCellValue("Ratings");
		    	//rowhead.createCell(9).setCellValue("Reviews");
		    	rowhead.createCell(7).setCellValue("Enrolment");
		    	rowhead.createCell(8).setCellValue("Level");
		    	rowhead.createCell(9).setCellValue("Type");
		    	rowhead.createCell(10).setCellValue("Duration");
		    	
		    	HSSFRow row;
		    	int i = 0;
		    	
		    	for(Course obj : courses) {
		    		row = sheet.createRow(++i);
		    		String Instructors ="";
		    		//String Departments ="";
		    		String Designation ="";
		    		ArrayList<Instructor>list = obj.instructors;
		    		for(int a =0;a<list.size();a++) {
		    			Instructor one = list.get(a);
		    			Instructors = Instructors+one.instructorName+", ";
		    			//Departments = one.instructorDepartment;
		    			Designation = Designation+one.instructorDesigination+", ";
		    		}
		    		//System.out.println(Instructors+"\n");
		    		//System.out.println(Departments+"\n");
		    		//System.out.println(Designation+"\n");
		    		
		    		row.createCell(0).setCellValue(i);  
			    	row.createCell(1).setCellValue(obj.courseName); 
			    	row.createCell(2).setCellValue(obj.url);
			    	row.createCell(3).setCellValue(Instructors);
			    	row.createCell(4).setCellValue(obj.offeredBy);
			    	row.createCell(5).setCellValue(Designation);
			    	//row.createCell(6).setCellValue(Departments);
			    	row.createCell(6).setCellValue(obj.skillsoffered);  
			    	//row.createCell(8).setCellValue(obj.ratings);
			    	//row.createCell(9).setCellValue(obj.reviews);
			    	row.createCell(7).setCellValue(obj.noOfEnrollments);
			    	row.createCell(8).setCellValue(obj.level);
			    	row.createCell(9).setCellValue(obj.courseType);
			    	row.createCell(10).setCellValue(obj.duration);
		    		
		    		
		    	}
		    	
		    	FileOutputStream fileOut;
				
					fileOut = new FileOutputStream(Filename);
					  
			    	wb.write(fileOut);  
			    	//closing the Stream  
			    	fileOut.close();  
			    	//closing the workbook  
			    	wb.close();  
			    	//prints the message on the console  
			    	System.out.println("Excel file has been generated successfully.");
				
				
			}
			
		}catch(Exception e) {
			
			e.printStackTrace();
		}
	}
	
	
	public static void edxCourses(String topic, String path, int pageNum, int courseNum, int pageSize) {
		String url = "https://www.edx.org/search?q="+topic+"&tab=course&availability=Available+now";
		ArrayList<Course> courses = new ArrayList<Course>();
		Course course;
		Instructor instructor;
		ArrayList instructors;
		
		
		ChromeDriver driver = new ChromeDriver(); //Open Chrome
		try {
			
		
		
		driver.get(url); //URL to open
		Thread.sleep(5000);
		
		
		if(!driver.findElements(By.className("pgn__modal-content-container")).isEmpty()) {
			driver.findElement(By.className("pgn__modal-close-container")).click();
			
		}
		
		Thread.sleep(5000);
		
		if(pageNum != 0) {
			for(int x =0; x<pageNum; x++) {
				driver.findElement(By.className("next")).click();
				Thread.sleep(5000);
			}
		}
		WebElement resultCount = driver.findElement(By.className("result-count"));
		String[] count = resultCount.getText().split(" "); 
		int resultTotal =  Integer.parseInt(count[0]);
		
		List<WebElement> results = driver.findElements(By.className("discovery-card"));
		int pages = resultTotal/results.size();
		if(pages>pageSize) {
			pages =pageSize;
		}
		

		for(int page = pageNum; page<pages; page++) {
			results = driver.findElements(By.className("discovery-card"));
			for(int i = courseNum;i<results.size();i++) {
				course = new Course();
				instructors = new ArrayList<Instructor>();
				results = driver.findElements(By.className("discovery-card"));
				WebElement element = results.get(i);
				
				String details = element.getText().toString();
				String[] courseDetails = details.split("\n");
				if(courseDetails.length > 4) {
					course.courseName = courseDetails[0];
					for(int j = 1;j < courseDetails.length; j++) {
						if(courseDetails[j].equals("Schools and Partners:")) {
							course.courseType = courseDetails[courseDetails.length-1];
							course.offeredBy = courseDetails[++j];
							
							break;
						}
						else {
							course.courseName = course.courseName+" "+ courseDetails[j];
							
						}
						
					}
					System.out.print("Course Name:"+course.courseName+"\n" );
					
					
					
				}
				else {
					course.courseName = courseDetails[0];
					course.courseType = courseDetails[3];
					course.offeredBy = courseDetails[2];
					
				}
				
				
				System.out.print("Page Number = "+page+"\t \t Course Num = "+i+"\n \n");
				//driver.findElement(By.className("d-card-wrapper")).click();
				//WebElement link = element.findElement(By.className("discovery-card-link"));
				//driver.findElement(By.xpath("//*[@id=\"main-content\"]/div/div[4]/div/div/div/div[2]/div/div[1]/div/a/div/div[3]/h3/span/span[1]/span")).click();
				
				if(!driver.findElements(By.className("edx-cookie-banner")).isEmpty()) {
					driver.findElement(By.className("close")).click();
					
				}
				
				//driver.findElement(By.className("discovery-card-link")).click();
				element.click();
				//Actions action = new Actions(driver);
				//action.moveToElement(element).click().perform();
				
				Thread.sleep(10000);
				
				//driver.findElement(By.xpath("//*[@id=\"main-content\"]/div/div[4]/div/div/div/div[2]/div/div[1]/div/a/div/div[3]/h3/span/span[1]/span")).click();
				
				course.url = driver.getCurrentUrl();
				List<WebElement> test = driver.findElements(By.className("at-a-glance"));
				Thread.sleep(2000);
				for(int c = 0;c<test.size(); c++ ) {
					String[] testString = test.get(c).getText().split("\n");
					course.level = testString[2];
					
				}
				
				test = driver.findElements(By.className("preview-expand-component"));
				for(int j = 0;j<test.size(); j++ ) {
					WebElement testing = test.get(j);
					if(testing.getText().contains("What you'll learn")) {
						String[] testString = testing.getText().split("\n");
						//	course.level = testString[2];
						//System.out.print(testString.length+"\n");
						course.skillsoffered = testString[2];
						for (int a = 3; a<testString.length-1; a++) {
							course.skillsoffered = course.skillsoffered +" , " + testString[a]; 
						}
						break;
					}
					
				}
				test = driver.findElements(By.className("course-selection"));
				String [] testDetails = test.get(1).getText().split("\n");
				//details = testing.getText();
				//System.out.print(testDetails[1]);
				course.noOfEnrollments = testDetails[1];
				test = driver.findElements(By.className("instructor-card"));
				for(int k = 0;k<test.size(); k++ ) {
					WebElement testing = test.get(k);
					String detail = testing.getText();
					String[] instructorDetails = detail.split("•"); 
					
					instructor = new Instructor();
					String department = "";
					if(instructorDetails.length>1) {
						String[] name = instructorDetails[0].split("\n");
						instructor.instructorName = name[0];
						department = name[1];
						if(name.length>2) {
							for(int b = 2; b<name.length;b++) {
								department = department +" "+ name[b]; 
							}
						}
						
						String [] desigination = instructorDetails[1].split("\n");
						department = department +" "+ desigination[0];
						if(desigination.length > 1) {
							for(int b = 1; b<desigination.length; b++) {
								department = department +", "+ desigination[b]; 
							}
						}
						
					}
					else {
						String[] name = instructorDetails[0].split("\n");
						if(name.length ==1) {
							instructor.instructorName = name[0];
							
						}
						else {
							instructor.instructorName = name[0];
							department = name[1];
							if(name.length>2) {
								for(int b = 2; b<name.length;b++) {
									department = department +", "+ name[b]; 
								}
							}
							
						}
						
						
						
					}
					
					
					instructor.instructorDesigination = department;
					
					instructors.add(instructor);
					course.instructors = instructors;
					
					
					
					
					
					
					
					
				}
				courses.add(course);
				driver.navigate().back();
				Thread.sleep(5000);
				//driver.navigate().refresh();
				//Thread.sleep(5000);
				
				//System.out.print(element.getText() + "\n \n");
				
			}
			if(courseNum != 0) {
				courseNum = 0;
			}
			if(!driver.findElements(By.className("next")).isEmpty()) {
				driver.findElement(By.className("next")).click();
			}
			Thread.sleep(5000);
			
			
			
		}
		
		
		
		driver.close();
		edxCourseFile(courses, path,topic);
		}catch (Exception ex) {

			edxCourseFile(courses, path,topic);
			ex.printStackTrace();
		}
		
		
	}
	
	public static void courseraCourseFile(ArrayList<Course>courses, String path, String topic) {
		try {
			String FileName = path+"Coursera-"+topic+".xlsx";
			File file = new File(FileName);
			HSSFWorkbook wb;

			
			if(file.exists()) {
				FileInputStream inputStream = new FileInputStream(file);
				wb = new HSSFWorkbook(inputStream);
				//System.out.print("File is writeable\n");
				HSSFSheet sheet = wb.getSheet(topic);
				int rowcount = sheet.getLastRowNum();
				
				
				HSSFRow row;
		    	int i = rowcount;
		    	
		    	for(Course obj : courses) {
		    		row = sheet.createRow(++i);
		    		String Instructors ="";
		    		//String Departments ="";
		    		String Designation ="";
		    		String Departments = "";
		    		ArrayList<Instructor>list = obj.instructors;
		    		for(int a =0;a<list.size();a++) {
		    			Instructor one = list.get(a);
		    			Instructors = Instructors+one.instructorName+", ";
		    			//Departments = one.instructorDepartment;
		    			Designation = Designation+one.instructorDesigination+", ";
		    			Departments = one.instructorDepartment;
		    		}
		    		//System.out.println(Instructors+"\n");
		    		//System.out.println(Departments+"\n");
		    		//System.out.println(Designation+"\n");
		    		
		    		row.createCell(0).setCellValue(i);  
			    	row.createCell(1).setCellValue(obj.courseName); 
			    	row.createCell(2).setCellValue(obj.url);
			    	row.createCell(3).setCellValue(Instructors);
			    	row.createCell(4).setCellValue(obj.offeredBy);
			    	row.createCell(5).setCellValue(Designation);
			    	row.createCell(6).setCellValue(Departments);
			    	row.createCell(7).setCellValue(obj.skillsoffered);  
			    	row.createCell(8).setCellValue(obj.ratings);
			    	row.createCell(9).setCellValue(obj.reviews);
			    	row.createCell(10).setCellValue(obj.noOfEnrollments);
			    	row.createCell(11).setCellValue(obj.level);
			    	row.createCell(12).setCellValue(obj.courseType);
			    	row.createCell(13).setCellValue(obj.duration);
		    		
		    		
		    	}
		    	inputStream.close();
				
				FileOutputStream fileOut = new FileOutputStream(FileName);
				wb.write(fileOut);
				fileOut.close();
				wb.close();
			}
			else {
				File file01 = new File(path);
				file01.mkdir();
				wb = new HSSFWorkbook();
		    	String Filename = path+"Coursera-"+topic+".xlsx";
		    	
		    	HSSFSheet sheet = wb.createSheet(topic);
		    	
		    	HSSFRow rowhead = sheet.createRow(0);
		    	
		    	rowhead.createCell(0).setCellValue("S.No.");  
		    	rowhead.createCell(1).setCellValue("Course Name"); 
		    	rowhead.createCell(2).setCellValue("URL");
		    	rowhead.createCell(3).setCellValue("Instructor");
		    	rowhead.createCell(4).setCellValue("Institute");
		    	rowhead.createCell(5).setCellValue("Designation");
		    	rowhead.createCell(6).setCellValue("Department");
		    	rowhead.createCell(7).setCellValue("Skills");  
		    	rowhead.createCell(8).setCellValue("Ratings");
		    	rowhead.createCell(9).setCellValue("Reviews");
		    	rowhead.createCell(10).setCellValue("Enrolment");
		    	rowhead.createCell(11).setCellValue("Level");
		    	rowhead.createCell(12).setCellValue("Type");
		    	rowhead.createCell(13).setCellValue("Duration");
		    	
		    	HSSFRow row;
		    	int i = 0;
		    	
		    	for(Course obj : courses) {
		    		row = sheet.createRow(++i);
		    		String Instructors ="";
		    		//String Departments ="";
		    		String Designation ="";
		    		String Departments = "";
		    		ArrayList<Instructor>list = obj.instructors;
		    		for(int a =0;a<list.size();a++) {
		    			Instructor one = list.get(a);
		    			Instructors = Instructors+one.instructorName+", ";
		    			//Departments = one.instructorDepartment;
		    			Designation = Designation+one.instructorDesigination+", ";
		    			Departments = one.instructorDepartment;
		    		}
		    		//System.out.println(Instructors+"\n");
		    		//System.out.println(Departments+"\n");
		    		//System.out.println(Designation+"\n");
		    		
		    		row.createCell(0).setCellValue(i);  
			    	row.createCell(1).setCellValue(obj.courseName); 
			    	row.createCell(2).setCellValue(obj.url);
			    	row.createCell(3).setCellValue(Instructors);
			    	row.createCell(4).setCellValue(obj.offeredBy);
			    	row.createCell(5).setCellValue(Designation);
			    	row.createCell(6).setCellValue(Departments);
			    	row.createCell(7).setCellValue(obj.skillsoffered);  
			    	row.createCell(8).setCellValue(obj.ratings);
			    	row.createCell(9).setCellValue(obj.reviews);
			    	row.createCell(10).setCellValue(obj.noOfEnrollments);
			    	row.createCell(11).setCellValue(obj.level);
			    	row.createCell(12).setCellValue(obj.courseType);
			    	row.createCell(13).setCellValue(obj.duration);
		    		
		    		
		    	}
		    	
		    	FileOutputStream fileOut;
				
					fileOut = new FileOutputStream(Filename);
					  
			    	wb.write(fileOut);  
			    	//closing the Stream  
			    	fileOut.close();  
			    	//closing the workbook  
			    	wb.close();  
			    	//prints the message on the console  
			    	System.out.println("Excel file has been generated successfully.");
				
				
			}
			
		}catch(Exception e) {
			
			e.printStackTrace();
		}
		
	}
	
	
	
	/*Coursera Courses*/
	
	public static void courseraCourses(String Topic,String Path,int pageNum, int courseNum, int pageSize) {
		ArrayList<Course> courses = new ArrayList<Course>();
		try {
			//ArrayList<Course> courses = new ArrayList<Course>();
			
			Course course;
			WebDriver driver = new ChromeDriver(); //Open Chrome			
//			ChromeDriver driver = new ChromeDriver(); //Open Chrome
			
			driver.get("https://www.coursera.org/"); //URL to open
			//driver.get("https://www.coursera.org/search?query=blockchain&page=5&index=prod_all_launched_products_term_optimization"); //URL to open
			//driver.manage().window().fullscreen();
			Thread.sleep(2000);
			
		    driver.findElement(By.className("react-autosuggest__input")).sendKeys( Topic, Keys.ENTER);
		    
			Thread.sleep(2000);
		
	    
		    List<WebElement> searches;
		    Thread.sleep(2000);
		    
		    if(pageNum!=0) {
		    	for(int x = 0;x < pageNum; x++) {
		    		Thread.sleep(2000);
		    		 List<WebElement> list = driver.findElements(By.className("css-16qv2i2"));
		 		    if(list.size() == 2) {
		 		    	list.get(1).click();
		 		    }
		 		    else {
		 		    	list.get(0).click();
		 		    }
		    		
		    	}
		    }
		    
		    String[] no_of_results = driver.findElement(By.className("rc-SearchResultsHeader")).getText().split(" ");
		    int TotalNumOfResults =Integer.parseInt(no_of_results[0]);
		    Thread.sleep(3000);
		    
		    int numOfResultsPerPage = driver.findElements(By.className("css-1pa69gt")).size(); //ais-InfiniteHits-item
		    int pages = TotalNumOfResults / numOfResultsPerPage;
		    
		    
		    if(pages > pageSize) {
		    	pages = pageSize;
		    	
		    }
		    		
	    
		for(int j = pageNum; j < pages; j++) {
	    	course = new Course();
	    	searches = driver.findElements(By.className("css-1pa69gt")); //ais-InfiniteHits-item
	    	Thread.sleep(3000);
	    	
		    if(searches.isEmpty()) {
		    	System.out.print("List is Empty \n");
		    	
		    }
		    else {
		    	//System.out.print(searches.size()+"\n");
		    	for(int i = courseNum; i < searches.size(); i++) {
		    		course = new Course();
		    		WebElement element = searches.get(i);
		    		//System.out.println(element.toString());
		    		String details =element.getText().toString();
		    		String[] courseDetails = details.split("\n");
		    		if(courseDetails.length ==7) {
		    			course.offeredBy = courseDetails[1];
		    			course.courseName = courseDetails[2];
		    			course.skillsoffered = courseDetails[3];
		    			course.ratings = courseDetails[4];
		    			course.reviews = courseDetails[5];
		    			String[] other = courseDetails[6].split("·");
	    				course.level = other[0];
		    			course.courseType = other[1];
		    			course.duration = other[2];	

		    			
		    		}
		    		else if(courseDetails.length ==6) {
		    			course.offeredBy = courseDetails[0];
		    			course.courseName = courseDetails[1];
		    			course.skillsoffered = courseDetails[2];
		    			course.ratings = courseDetails[3];
		    			course.reviews = courseDetails[4];
		    			String[] other = courseDetails[5].split("·");
		    			course.level = other[0];
		    			course.courseType = other[1];
		    			course.duration = other[2];
		    			
		    		}
		    		else {
		    			
		    			continue;
		    			
		    		}
		    		
		    		
		    		//System.out.print(details+"\n \n");
		    		element.click();
		    		
		    		Thread.sleep(2000);
					
		    		
		    		ArrayList<String> tabs2 = new ArrayList<String> (driver.getWindowHandles());
		    	    driver.switchTo().window(tabs2.get(1));
		    	    
		    	    course.url = driver.getCurrentUrl();
		    	    if(!driver.findElements(By.className("_1fpiay2")).isEmpty()) {
		    	    	course.noOfEnrollments = driver.findElement(By.className("_1fpiay2")).getText();
		    	    }
		    	    //driver.findElement(By.className("_1wb6qi0n")).findElement("_y1d9czk").;
		    	    
		    	    if(!driver.findElements(By.className("_y1d9czk")).isEmpty()){
		    	    	//!driver.findElements(By.className("_y1d9czk")).isEmpty()
		    	    	if(driver.findElement(By.className("_y1d9czk")).isDisplayed()) {
		    	    		
		    	    		driver.findElement(By.className("_y1d9czk")).click();
							Thread.sleep(2000);
							
							List<WebElement> instructorDetails = driver.findElements(By.className("instructor-wrapper"));
				    	    
				    	    for(int a=0; a<instructorDetails.size();a++) {
				    	    	String[] InstructorDetails = instructorDetails.get(a).getText().split("\n");
					    	    Instructor instructor1 = new Instructor();
				    	    	instructor1.instructorName = InstructorDetails[0];
					    	    instructor1.instructorDesigination = InstructorDetails[1];
					    	    instructor1.instructorDepartment = InstructorDetails[2];
					    	    course.instructors.add(instructor1);
				    	    	
				    	    	}
		    	    		}
		    	    	}
		    	    driver.close();
		    	    driver.switchTo().window(tabs2.get(0));
		    	    courses.add(course);
		    	    System.out.print("Course Name: "+ course.courseName+"\n");
		    	    System.out.print("Page Number: "+ j +"\t \t Course Number: "+ i +"\n \n");
		    	    
				    
		    	
		    	}
		    	if(courseNum!=0) {
		    		courseNum = 0;
		    	}
		    
			    
		    }
		    searches.clear();
		    List<WebElement> list = driver.findElements(By.className("css-16qv2i2"));
		    if(list.size() == 2) {
		    	list.get(1).click();
		    }
		    else {
		    	list.get(0).click();
		    }
		    
			Thread.sleep(2000);
			
	    }
	    
	    
	    driver.close();
	    
	    System.out.println("Script End."+ "\n");
	    
	    courseraCourseFile(courses, Path, Topic);	
	    HSSFWorkbook wb = new HSSFWorkbook();
	    	String Filename = Path+"Coursera-"+Topic+".xlsx";
	    	
	    	HSSFSheet sheet = wb.createSheet(Topic);
	    	
	    	HSSFRow rowhead = sheet.createRow(0);
	    	
	    	rowhead.createCell(0).setCellValue("S.No.");  
	    	rowhead.createCell(1).setCellValue("Course Name"); 
	    	rowhead.createCell(2).setCellValue("URL");
	    	rowhead.createCell(3).setCellValue("Instructor");
	    	rowhead.createCell(4).setCellValue("Institute");
	    	rowhead.createCell(5).setCellValue("Designation");
	    	rowhead.createCell(6).setCellValue("Department");
	    	rowhead.createCell(7).setCellValue("Skills");  
	    	rowhead.createCell(8).setCellValue("Ratings");
	    	rowhead.createCell(9).setCellValue("Reviews");
	    	rowhead.createCell(10).setCellValue("Enrolment");
	    	rowhead.createCell(11).setCellValue("Level");
	    	rowhead.createCell(12).setCellValue("Type");
	    	rowhead.createCell(13).setCellValue("Duration");
	    	
	    	HSSFRow row;
	    	int i = 0;
	    	
	    	for(Course obj : courses) {
	    		row = sheet.createRow(++i);
	    		String Instructors ="";
	    		String Departments ="";
	    		String Designation ="";
	    		ArrayList<Instructor>list = obj.instructors;
	    		for(int a =0;a<list.size();a++) {
	    			Instructor one = list.get(a);
	    			Instructors = Instructors+one.instructorName+", ";
	    			Departments = one.instructorDepartment;
	    			Designation = Designation+one.instructorDesigination+", ";
	    		}
	    		System.out.println(Instructors+"\n");
	    		System.out.println(Departments+"\n");
	    		System.out.println(Designation+"\n");
	    		
	    		row.createCell(0).setCellValue(i);  
		    	row.createCell(1).setCellValue(obj.courseName); 
		    	row.createCell(2).setCellValue(obj.url);
		    	row.createCell(3).setCellValue(Instructors);
		    	row.createCell(4).setCellValue(obj.offeredBy);
		    	row.createCell(5).setCellValue(Designation);
		    	row.createCell(6).setCellValue(Departments);
		    	row.createCell(7).setCellValue(obj.skillsoffered);  
		    	row.createCell(8).setCellValue(obj.ratings);
		    	row.createCell(9).setCellValue(obj.reviews);
		    	row.createCell(10).setCellValue(obj.noOfEnrollments);
		    	row.createCell(11).setCellValue(obj.level);
		    	row.createCell(12).setCellValue(obj.courseType);
		    	row.createCell(13).setCellValue(obj.duration);
	    		
	    		
	    	}
	    	FileOutputStream fileOut = new FileOutputStream(Filename);  
	    	wb.write(fileOut);  
	    	//closing the Stream  
	    	fileOut.close();  
	    	//closing the workbook  
	    	wb.close();  
	    	//prints the message on the console  
	    	System.out.println("Excel file has been generated successfully.");  
	    
	    
		}
	    catch (Exception e) {
			// TODO Auto-generated catch block
	    	courseraCourseFile(courses, Path,Topic);
			e.printStackTrace();
		}
	}
	
	    

}
