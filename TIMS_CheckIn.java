package Utilities;

import java.io.File;
import java.io.FileInputStream;
import java.util.Properties;

import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPasswordField;
import javax.swing.JTextField;

import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.sikuli.script.Screen;

public class TIMS_CheckIn {

	public static final Properties PROPERTIES = Utilities.PROPERTIES;
	static String SIKULI_IMAGES = PROPERTIES.getProperty("ImagesPath");
	static String TEST_DATA_FILEPATH = PROPERTIES.getProperty("TestDataPath");

	public static void main(String[] args) throws Exception {
		Screen sobj = new Screen ();
		//OLD -> File file = new File("C:\\DEV\\VerintAutomation\\Verint_Shakedown_Automation\\files\\IEDriverServer.exe");
		String ieDriverFilepath = PROPERTIES.getProperty("IEDriverServerPath");
		File file = new File(ieDriverFilepath);
		System.setProperty("webdriver.ie.driver", file.getAbsolutePath());
		DesiredCapabilities capabilities = DesiredCapabilities.internetExplorer();
		capabilities.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS, true);
		WebDriver driver = new InternetExplorerDriver(capabilities);
		JLabel jUserName = new JLabel("User Name");
        JTextField userName = new JTextField();
        JLabel jPassword = new JLabel("Password");
        JTextField password = new JPasswordField();
        Object[] ob = {jUserName, userName, jPassword, password};
        int result = JOptionPane.showConfirmDialog(null, ob, "Please Enter User Id and Password", JOptionPane.OK_CANCEL_OPTION);
 
        if (result == JOptionPane.OK_OPTION) {
            String userNameValue = userName.getText().toUpperCase();
            String passwordValue = password.getText();
            driver.get("https://sftims.opr.statefarm.org/TIMS/Pages/MyTestIDs/ViewMyIDs.aspx");
    		Thread.sleep(3000);
			if (sobj.exists(SIKULI_IMAGES + "\\WindowsSecurity_UserName.png") != null) {
				sobj.type(SIKULI_IMAGES + "\\WindowsSecurity_UserName.png","opr\\"+userNameValue);
				Thread.sleep(1000);			
				if (sobj.exists(SIKULI_IMAGES + "\\WindowsSecurity_Password.png") != null) {
					sobj.type(SIKULI_IMAGES + "\\WindowsSecurity_Password.png", passwordValue);
					Thread.sleep(1000);
				}
				if (sobj.exists(SIKULI_IMAGES + "\\WindowsSecurity_OK.png") != null) {
					sobj.click(SIKULI_IMAGES + "\\WindowsSecurity_OK.png");
					Thread.sleep(1000);
				}
				if (sobj.exists(SIKULI_IMAGES + "\\WindowsSecurity_OK1.png") != null) {
					sobj.click(SIKULI_IMAGES + "\\WindowsSecurity_OK1.png");
					Thread.sleep(1000);
				}
			}
			String userID;
			String testID;
			FileInputStream fig = new FileInputStream(TEST_DATA_FILEPATH);
			Workbook workbook = Workbook.getWorkbook(fig);
		    Sheet worksheet = workbook.getSheet("TestIds");
		    int rows = worksheet.getRows();
		    for (int i=1; i<rows; i++) {
				userID = worksheet.getCell(0, i).getContents();
				String id;
				id = userID.trim().toUpperCase();
				driver.manage().window().maximize();
				driver.findElement(By.id("MainContent_IDGrid1_txtFilter")).clear();
				driver.findElement(By.id("MainContent_IDGrid1_txtFilter")).sendKeys(id);
				Thread.sleep(3000);

				if (driver.findElement(By.xpath("//div[@id='alertText']")).getText().contains("No results found!")) {
					if (driver.findElements(By.xpath("//input[@id='btnAlertClose']")).size() != 0) {
						driver.findElement(By.xpath("//input[@id='btnAlertClose']")).click();
						System.out.println("No Results found");
						break;
					}
				}

				if (driver.findElements(By.xpath("//table[@id='MainContent_IDGrid1_gv']")).size() != 0) {
					int rc = driver.findElements(By.xpath("//table[@id='MainContent_IDGrid1_gv']/tbody/tr")).size();
					if (rc > 1) {
						if (driver.findElement(By.xpath("//table[@id='MainContent_IDGrid1_gv']/tbody/tr[2]/td[10]")).getText().trim().contains("Checked Out by " + userNameValue)) {
							testID = driver.findElement(By.xpath("//table[@id='MainContent_IDGrid1_gv']/tbody/tr[2]/td[3]")).getText().trim().toUpperCase();
							if (testID.contains(id)) {
								driver.findElement(By.xpath("//table[@id='MainContent_IDGrid1_gv']/tbody/tr[2]/td[2]/input[@name='RadioGroup'][@type='radio']")).click();
								Thread.sleep(1000);
								if (driver.findElements(By.xpath("//input[@id='MainContent_IDGrid1_btnCheckIn']")).size() != 0) {
									driver.findElement(By.xpath("//input[@id='MainContent_IDGrid1_btnCheckIn']")).click();
									Thread.sleep(10000);
								}
								if (driver.findElements(By.xpath("//div[@id='alertButton']/input[@id='btnAlertClose']")).size() != 0) {
									driver.findElement(By.xpath("//div[@id='alertButton']/input[@id='btnAlertClose']")).click();
								}
							}
						}
						if (driver.findElement(By.xpath("//table[@id='MainContent_IDGrid1_gv']/tbody/tr[2]/td[10]")).getText().contains("Available")) {
							System.out.println("Available. Already Checked In");
						}
						if (!driver.findElement(By.xpath("//table[@id='MainContent_IDGrid1_gv']/tbody/tr[2]/td[10]")).getText().contains("Checked Out by " + userNameValue)) {
							System.out.println(driver.findElement(By.xpath("//table[@id='MainContent_IDGrid1_gv']/tbody/tr[2]/td[10]")).getText() + " different User.Cannot be Checked In.");
						}
					}
				}
				deletePassword(driver, "", 1, i);
			}
		//If OK dialog
		workbook.close();
		fig.close();
		driver.close();
		driver.quit();
		}
	}

	public static void deletePassword(WebDriver driver, String pwd, int col, int row) throws Exception {
		Workbook workbook = Workbook.getWorkbook(new File(TEST_DATA_FILEPATH));
	    WritableWorkbook writableWorkbook = Workbook.createWorkbook(new File(TEST_DATA_FILEPATH), workbook);
	    WritableSheet sheet = writableWorkbook.getSheet("TestIds");
	    WritableCell cell;
	    Label label = new Label(col, row, pwd);
	    cell = (WritableCell) label;
	    sheet.addCell(cell);
	    writableWorkbook.write();
	    writableWorkbook.close();
	}
}
