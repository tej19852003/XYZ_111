package Utilities;

import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;
import java.util.Set;
import java.lang.reflect.Method;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import org.openqa.selenium.os.WindowsUtils;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.sikuli.script.Screen;

import ScreenObjects.LoginScreen;
import ScreenObjects.VerintHomePageScreen;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.format.Pattern;
import jxl.format.UnderlineStyle;

import jxl.write.Label;

import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableHyperlink;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class Utilities {

    public static ExtentReports extent = ExtentReports.get(Utilities.class);
    public static Screen sobj = new Screen();
    public static String datapath;
    public static String sheetname;
	public static String frameTreeViewLeftPane = "oLeftPaneContent";
	public static String frameRightPaneContent = "oRightPaneContent";

    public static Properties PROPERTIES = new Properties();
    public static File PROPERTIES_FILEPATH = new File("..\\bin\\Properties\\global-variables.properties");

	public static void loadGlobalLocators() throws Exception {
		PROPERTIES.load(new FileInputStream(PROPERTIES_FILEPATH));
	}
	
	public static boolean selectMenu(WebDriver driver, String username, String password, String tabName, String menuItem) throws Exception {
		boolean flag = false;
		LoginScreen.setTextInUsername(driver, username);
		LoginScreen.setTextInPassword(driver, password);
		LoginScreen.clickLogin(driver);
		if (VerintHomePageScreen.verifyVerintHomePage(driver)) {
			VerintHomePageScreen.selectMenuItem(driver, tabName, menuItem);
			Thread.sleep(6000);
			if (driver.findElements(By.linkText(tabName)).size() == 0) {
				Utilities.logout(driver);  //logout
			} else {
				return flag = true;
			}
		} else {
			return flag = false;
		}
		return flag;
	}
	
	public static String getPassword(WebDriver driver, int col, int row) throws Exception {
		String password = "";
        String excelFilepath = PROPERTIES.getProperty("TestDataPath");
        FileInputStream stream = new FileInputStream(excelFilepath);
        Workbook workbook = Workbook.getWorkbook(stream);
	    Sheet worksheet = workbook.getSheet("TestIds");
	    password = worksheet.getCell(col, row).getContents();
	    return password;
    }
	
	public static String getUserID(WebDriver driver, int col, int row) throws Exception {
		String id = "";
        String excelFilepath = PROPERTIES.getProperty("TestDataPath");
        FileInputStream fis = new FileInputStream(excelFilepath);
        Workbook workbook = Workbook.getWorkbook(fis);
	    Sheet worksheet = workbook.getSheet("TestIds");
	    id = worksheet.getCell(col, row).getContents();
	    workbook.close();
	    fis.close();
	    return id;
	 }
	
	public static void windowsSecurityCredentials(WebDriver driver, String userName, String password) throws Exception {
		Thread.sleep(3000);
		if (sobj.exists(PROPERTIES.getProperty("ImagesPath") + "\\WindowsSecurity_UserName.png") != null) {
			sobj.type(PROPERTIES.getProperty("ImagesPath") + "\\WindowsSecurity_UserName.png", userName);
			if (sobj.exists(PROPERTIES.getProperty("ImagesPath") + "\\WindowsSecurity_Password.png") != null) {
				sobj.type(PROPERTIES.getProperty("ImagesPath") + "\\WindowsSecurity_Password.png", password);
			}
			if (sobj.exists(PROPERTIES.getProperty("ImagesPath") + "\\WindowsSecurity_OK.png") != null) {
				sobj.click(PROPERTIES.getProperty("ImagesPath") + "\\WindowsSecurity_OK.png");
			}
		}
	}

	public static boolean searchItem(WebDriver driver, String searchText) throws Exception {
		boolean flag = false;
		boolean temp = false;
		Robot r = new Robot();
		
		for (int s=1; s<=5; s++) {
			r.keyPress(KeyEvent.VK_CONTROL);
			r.keyPress(KeyEvent.VK_F);
			Thread.sleep(1000);
			r.keyRelease(KeyEvent.VK_CONTROL);
			r.keyRelease(KeyEvent.VK_F);
			Thread.sleep(3000);
			if (sobj.exists(PROPERTIES.getProperty("ImagesPath") + "\\FindText.png") != null) {
				sobj.type(PROPERTIES.getProperty("ImagesPath") + "\\FindText.png", searchText);
				Thread.sleep(5000);
				if (sobj.exists(PROPERTIES.getProperty("ImagesPath") + "\\NoMatchesFound.png") != null) {
					temp = false;
				} else {
					temp = true;
				}
				System.out.println(temp);
				Utilities.sikuliClick(driver, PROPERTIES.getProperty("ImagesPath") + "\\FindText_Close.png");
				break;
			}
		}
		if (temp == true) {
			extent.log(LogStatus.PASS, "Search Item:" + searchText + " successful");
			flag = true;
		} else {
			extent.log(LogStatus.FAIL, "Search Item:" + searchText + " NOT successful");
			flag = false;
		}	
		return flag;
	}

	public static boolean sikuliClick(WebDriver driver, String imagePath) throws Exception {
		boolean flag = false;
		if (sobj.exists(imagePath) != null) {
			sobj.click(imagePath);
			Thread.sleep(3000);
			flag = true;
		}
		return flag;
	}

	public static void sikuliType(WebDriver driver, String imagePath, String text) throws Exception {
		if (sobj.exists(imagePath) != null) {
			sobj.type(imagePath, text);
			Thread.sleep(2000);
		}
	}
	
	public static void sikuliRightClick(WebDriver driver, String imagePath) throws Exception {
		if (sobj.exists(imagePath) != null) {
			sobj.rightClick(imagePath);
			System.out.println("clicked on this image");
			Thread.sleep(3000);
		}
	}
	
	public static boolean sikuliExist(WebDriver driver, String imagePath) throws Exception {
		boolean flag = true;
		if (sobj.exists(imagePath) != null) {
			extent.log(LogStatus.PASS, "Call is in Playback state");
			extent.log(LogStatus.PASS, "", "", Utilities.captureScreenShot(driver, "QM05_07_14_Playback_EvalForm_RemarksBy"));
		}
		else {
			extent.log(LogStatus.FAIL, "Call is NOT in Playback state");
			extent.log(LogStatus.FAIL, "", "", Utilities.captureScreenShot(driver, "QM05_07_14_Playback_EvalForm_RemarksBy"));
			flag = false;
		}
		return flag;
	}

	public static void switchFrame(WebDriver driver, String data) {
        List<WebElement> iframeList = driver.findElements(By.tagName("iframe"));
        List<WebElement> frameList = driver.findElements(By.tagName("frame"));
        frameList.addAll(iframeList);
        System.out.println("Total frames present - " + frameList.size());
        for (WebElement frame : frameList) {
            if ((frame.getAttribute("id").equals(data)) || (frame.getAttribute("name").equals(data))) {
                System.out.println("Switching on frame with ID - " + frame.getAttribute("id"));
                System.out.println("*** Name*** - " + frame.getAttribute("name"));
                driver.switchTo().frame(frame);
                break;
            }
        }
    }

	public static void selectRightPaneView(WebDriver driver) throws Exception {
		driver.switchTo().defaultContent();	
		Thread.sleep(2000);
		WebElement oRightPaneContentFrame = (new WebDriverWait(driver, 10)).until(ExpectedConditions.elementToBeClickable(By.id(frameRightPaneContent)));
		driver.switchTo().frame(oRightPaneContentFrame);
		Thread.sleep(3000);
	}

	public static void selectLeftTreeFrame(WebDriver driver) throws Exception {
		driver.switchTo().defaultContent();
		Thread.sleep(1000);
		WebElement leftTreeContentFrame1 = (new WebDriverWait(driver, 10)).until(ExpectedConditions.elementToBeClickable(By.id(frameTreeViewLeftPane)));
		driver.switchTo().frame(leftTreeContentFrame1);
		Thread.sleep(1000);
	}
	
	public static void logout(WebDriver driver) throws Exception {
		driver.switchTo().defaultContent();
		if (driver.findElements(By.id("utilityPanePC_LOGOUT_spn_id")).size() != 0)
			driver.findElement(By.id("utilityPanePC_LOGOUT_spn_id")).click();
		Thread.sleep(3000);
	}

	public static String setWindowFocus(WebDriver driver) throws Exception {
		Set<String> windowIds = driver.getWindowHandles();
		Iterator<String> itererator = windowIds.iterator(); 			
		String mainWinID = itererator.next();   //main window
		Thread.sleep(2000);
		String  popWindow = itererator.next();  //popup window
		driver.switchTo().window(popWindow);    //Switch the driver to the popup window
		return mainWinID;
	}
	
	public static boolean folderExist(String path, String type) {
		boolean flag = true;
		File prodDir = new File(path);
		boolean exists = prodDir.exists();
	    if (exists) {
	    	extent.log(LogStatus.PASS, path + " " + type + " exists");
	    } else {
	    	extent.log(LogStatus.FAIL, path + " " + type + " does not exists");
	    	return flag == false;
	    }
	    return flag;
	}

	public static void displayDirectoryContents(File directory) {
	    try {
	        File[] files = directory.listFiles();
	        for (File file : files) {
	            if (file.isDirectory()) {
	            	extent.log(LogStatus.INFO, "Directory Name=>:" + file.getCanonicalPath());
	                displayDirectoryContents(file);
	            } else {
	            	extent.log(LogStatus.INFO, "file Name=>" + file.getCanonicalPath());
	            }
	        }
	    } catch (IOException e) {
	        e.printStackTrace();
	    }
	}
	
	public static void waitForPageLoad(WebDriver driver, By str) throws Exception {
		for (int i=0; i<15 && driver.findElements(str).size() == 0; i++) {
           Thread.sleep(1000);
        }
	}
	
	public static void testCaseSetup(String htmlReportName, String testCaseName) throws Exception {
		Runtime.getRuntime().exec("taskkill /F /IM IEDriverServer.exe");
		Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe");
		String screenshotDir = PROPERTIES.getProperty("ScreenShotPath");
		new File(screenshotDir).mkdirs();

		String reportsDir = PROPERTIES.getProperty("ReportPath");
		new File(reportsDir).mkdirs();		
        
		ExtentReports extent = ExtentReports.get(Utilities.class);		
		extent.init(PROPERTIES.getProperty("ReportPath") + "\\" + htmlReportName + ".html", true);
		extent.config().documentTitle("Verint Automation Test Result Report");        
        extent.config().reportHeadline(testCaseName + " Result Report");
        extent.config().displayCallerClass(false);
        extent.config().useExtentFooter(false);
        extent.startTest(testCaseName);
	}

    public static String captureScreenShot(WebDriver driver, String pageName) throws Exception {
        String screenshotPath = "";
        File scrFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
        try {
            //Copy file object to designated location
            screenshotPath = PROPERTIES.getProperty("ScreenShotPath") + pageName + "_" + System.currentTimeMillis() + ".png";
            FileUtils.copyFile(scrFile, new File(screenshotPath));
            Thread.sleep(2000);
        } catch (IOException e) {
            System.out.println("Error while generating screenshot:\n" + e.toString());
        }
        return screenshotPath;
    }

    public static void verintExecution(String name) throws Exception {
        Utilities.loadGlobalLocators();

        datapath = PROPERTIES.getProperty("TestDataPath");
        if (name.contains("BO")) {
            sheetname = "BO_TestSet";
        }
        if (name.contains("DPA")) {
            sheetname = "DPA_TestSet";
        }
        if (name.contains("QM")) {
            sheetname = "QM_TestSet";
        }
        else if (name.contains("WFM")) {
            sheetname = "WFM_TestSet";
        }

        String run = "";
        String strOne = "";
        FileInputStream fis = new FileInputStream(datapath);
        Workbook workbook = Workbook.getWorkbook(fis);
        Sheet worksheet = workbook.getSheet(sheetname);

        for (int row = 1; row < worksheet.getRows(); row++) {
            run = workbook.getSheet(sheetname).getCell(3, row).getContents();
            if (run.equals("Y") ||  run.equals("y")) {
                strOne = worksheet.getCell(1, row).getContents();
                Class cls = Class.forName(strOne);
				System.out.println("TestScriptName: " + strOne);
				System.out.println("ClassName: " + cls);
                String methodName = worksheet.getCell(2, row).getContents();
				System.out.println("Start of Test case execution: " + methodName);
				Method method = cls.getDeclaredMethod(methodName);
                method.invoke(cls.getClass());
                System.out.println("End of Test case execution: " + methodName);
                System.out.println("####################################################################");
            }
        }
        workbook.close();
        fis.close();
     }

	public static void verintScriptStatus(boolean status, String name, String htmlReportName, int col, int row) throws Exception {
        datapath = PROPERTIES.getProperty("TestDataPath");

        if (name.contains("BO")) {
			sheetname = "BO_TestSet";
        }
		if (name.contains("DPA")) {
			sheetname = "DPA_TestSet";
        }
		if (name.contains("QM")) {
			sheetname = "QM_TestSet";
        }
		else if (name.contains("WFM")) {
			sheetname = "WFM_TestSet";
        }

		Workbook workbook = Workbook.getWorkbook(new File(datapath));
	    WritableWorkbook writableWorkbook = Workbook.createWorkbook(new File(datapath), workbook);
	    WritableSheet sheet = writableWorkbook.getSheet(sheetname);
	    File file = new File(PROPERTIES.getProperty("ReportPath") + "\\" + htmlReportName + ".html");

        String scriptStatus = "";
	    if (status==true) {
	    	scriptStatus="PASS";
	    } else {
	    	scriptStatus="FAIL";
	    }
	     
        WritableHyperlink hyperlink = new WritableHyperlink(col, row, file, scriptStatus);
        sheet.addHyperlink(hyperlink);

        if (scriptStatus=="PASS") {
            Label label = new Label(col, row, "PASS", getCellFormat(Colour.BRIGHT_GREEN, Pattern.GRAY_50));
		    sheet.addCell(label);
        } else if (scriptStatus=="FAIL") {
			Label label = new Label(col, row, "FAIL", getCellFormat(Colour.ROSE, Pattern.GRAY_50));
		    sheet.addCell(label);
        }

	    writableWorkbook.write();
	    writableWorkbook.close();
	    Thread.sleep(3000);
	 }

    private static WritableCellFormat getCellFormat(jxl.format.Colour green, jxl.format.Pattern gray25) throws WriteException {
        WritableFont cellFont = new WritableFont(WritableFont.ARIAL, 9);
        cellFont.setBoldStyle(WritableFont.BOLD);
        cellFont.setUnderlineStyle(UnderlineStyle.SINGLE);
        WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
        cellFormat.setBackground(green, gray25);
        cellFormat.setAlignment(Alignment.CENTRE);
        return cellFormat;
    }
	 
    public static void killProcess() throws Exception {
        String line;
        Process p = Runtime.getRuntime().exec(System.getenv("windir") + "\\system32\\" + "tasklist.exe");
        BufferedReader input = new BufferedReader(new InputStreamReader(p.getInputStream()));
        while ((line = input.readLine()) != null) {
            if (line.contains("javaw.exe")) {
                WindowsUtils.tryToKillByName("javaw.exe");
            }
            if (line.contains("EXCEL.EXE")) {
                WindowsUtils.tryToKillByName("EXCEL.EXE");
            }
            if (line.contains("iexplore")) {
                WindowsUtils.tryToKillByName("iexplore.exe");
            }
        }
        input.close();
    }
}
