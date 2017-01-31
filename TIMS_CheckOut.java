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

public class TIMS_CheckOut {

	public static final Properties PROPERTIES = Utilities.PROPERTIES;
	static String SIKULI_IMAGES = PROPERTIES.getProperty("ImagesPath");
	static String TEST_DATA_FILEPATH = PROPERTIES.getProperty("TestDataPath");
	static String IE_DRIVER = PROPERTIES.getProperty("IEDriverServerPath");

	public static void main(String[] args) throws Exception {
	
		Screen sobj = new Screen ();
		String pwd = "";
		File file = new File(IE_DRIVER);
		System.setProperty("webdriver.ie.driver", file.getAbsolutePath());
		WebDriver driver;
		DesiredCapabilities capabilities = DesiredCapabilities.internetExplorer();
		capabilities.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS, true);
		driver = new InternetExplorerDriver(capabilities);

		JLabel jUserName = new JLabel("User Name");
        JTextField userName = new JTextField();
        JLabel jPassword = new JLabel("Password");
        JTextField password = new JPasswordField();
        Object[] ob = {jUserName, userName, jPassword, password};
        int result = JOptionPane.showConfirmDialog(null, ob, "Please Enter your SF alias Id and Password", JOptionPane.OK_CANCEL_OPTION);
 
        if (result == JOptionPane.OK_OPTION) {
            String userNameValue = userName.getText().toUpperCase();
            String passwordValue = password.getText();

            driver.get("https://sftims.opr.statefarm.org/TIMS/Pages/MyTestIDs/ViewMyIDs.aspx");
    		Thread.sleep(3000);
			if (sobj.exists(SIKULI_IMAGES + "\\WindowsSecurity_UserName.png") != null) {
				sobj.type(SIKULI_IMAGES + "\\WindowsSecurity_UserName.png","opr\\" + userNameValue);
				Thread.sleep(1000);
			}
			if (sobj.exists(SIKULI_IMAGES + "\\WindowsSecurity_Password.png") != null) {
				sobj.type(SIKULI_IMAGES + "\\WindowsSecurity_Password.png", passwordValue);
				Thread.sleep(1000);
			}
			if (sobj.exists(SIKULI_IMAGES + "\\WindowsSecurity_OK.png") != null) {
				sobj.click(SIKULI_IMAGES + "\\WindowsSecurity_OK.png");
				Thread.sleep(1000);
			}
			if (sobj.exists(SIKULI_IMAGES + "\\img\\WindowsSecurity_OK1.png") != null) {
				sobj.click(SIKULI_IMAGES + "\\img\\WindowsSecurity_OK1.png");
				Thread.sleep(1000);
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

				pageLoad(driver, id);

				if (driver.findElement(By.xpath("//div[@id='alertText']")).getText().contains("No results found!")) {
					break;
				}

				for (int j=1; j<10; j++) {
					Thread.sleep(2000);
					if (driver.findElement(By.xpath("//div[@id='MainContent_IDGrid1_SizeSection']/label")).getText().contains("Show")
                            || driver.findElement(By.xpath("//div[@id='alertText']")).getText().contains("No results found!")) {
						break;
					}
				}

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
						if (driver.findElement(By.xpath("//table[@id='MainContent_IDGrid1_gv']/tbody/tr[2]/td[10]")).getText().contains("Available")) {
							testID = driver.findElement(By.xpath("//table[@id='MainContent_IDGrid1_gv']/tbody/tr[2]/td[3]")).getText().trim().toUpperCase();
							if (testID.contains(id)) {
								driver.findElement(By.xpath("//table[@id='MainContent_IDGrid1_gv']/tbody/tr[2]/td[2]/input[@name='RadioGroup'][@type='radio']")).click();
								Thread.sleep(2000);
								if (driver.findElements(By.xpath("//input[@id='MainContent_IDGrid1_btnCheckOut']")).size() !=0) {
									driver.findElement(By.xpath("//input[@id='MainContent_IDGrid1_btnCheckOut']")).click();
								}
								Thread.sleep(2000);
								if (driver.findElements(By.xpath("//div[@id='MainContent_IDGrid1_pwText']")).size() !=0) {
									pwd=driver.findElement(By.xpath("//div[@id='MainContent_IDGrid1_pwText']")).getText().trim();
								}
								Thread.sleep(2000);
								if (driver.findElements(By.xpath("//input[@id='MainContent_IDGrid1_btnPWClose']")).size() != 0) {
									driver.findElement(By.xpath("//input[@id='MainContent_IDGrid1_btnPWClose']")).click();
								}
							}
						}
						//Checked out by same user id
						else if (driver.findElement(By.xpath("//table[@id='MainContent_IDGrid1_gv']/tbody/tr[2]/td[10]")).getText().trim().contains("Checked Out by "+userNameValue)) {
							testID = driver.findElement(By.xpath("//table[@id='MainContent_IDGrid1_gv']/tbody/tr[2]/td[3]")).getText().trim().toUpperCase();
							if (testID.contains(id)) {
								driver.findElement(By.xpath("//table[@id='MainContent_IDGrid1_gv']/tbody/tr[2]/td[2]/input[@name='RadioGroup'][@type='radio']")).click();
								Thread.sleep(1000);
								if (driver.findElements(By.xpath("//input[@id='MainContent_IDGrid1_btnViewPW']")).size() !=0) {
									driver.findElement(By.xpath("//input[@id='MainContent_IDGrid1_btnViewPW']")).click();
								}
								if (driver.findElements(By.xpath("//div[@id='MainContent_IDGrid1_pwText']")).size() !=0) {
									pwd=driver.findElement(By.xpath("//div[@id='MainContent_IDGrid1_pwText']")).getText().trim();
								}
								Thread.sleep(2000);
								if (driver.findElements(By.xpath("//input[@id='MainContent_IDGrid1_btnPWClose']")).size() !=0) {
									driver.findElement(By.xpath("//input[@id='MainContent_IDGrid1_btnPWClose']")).click();
								}
							}
						}
						else if (!driver.findElement(By.xpath("//table[@id='MainContent_IDGrid1_gv']/tbody/tr[2]/td[10]")).getText().contains("Checked Out by " + userNameValue)) {
							System.out.println(driver.findElement(By.xpath("//table[@id='MainContent_IDGrid1_gv']/tbody/tr[2]/td[10]")).getText() + " different User.Try Again.");
						}
					}
				}
				writePassword(driver, pwd, 1, i);
			}
		   
			//If OK dialog
		    workbook.close();
		    fig.close();

			//Remove all status from excel sheets
		    deleteStatusFromExcel();
		    
		    driver.close();
		    driver.quit();
		    System.out.println("End of TIMS Checkout");
	    	System.out.println("####################################################################");
        }
    }
	
	public static void pageLoad(WebDriver driver, String id) throws Exception {
		for (int i=1; i<=40; i++) {
			if (driver.findElement(By.xpath("//table[@id='MainContent_IDGrid1_gv']/tbody/tr[2]/td[3]")).getText().trim().toUpperCase().contains(id)) {
				break;
			} else {
				Thread.sleep(2000);
			}
		}
	}

	public static void deleteStatusFromExcel() throws Exception {
		Workbook workbook = Workbook.getWorkbook(new File(TEST_DATA_FILEPATH));
	    WritableWorkbook writableWorkbook = Workbook.createWorkbook(new File(TEST_DATA_FILEPATH), workbook);

	    WritableSheet sheet = writableWorkbook.getSheet("BO_TestSet");
	    WritableCell cell;	    
	    int row_bo = sheet.getRows();
	    for (int i=1; i<row_bo; i++) {
		    Label label = new Label(4, i, "");
		    cell = (WritableCell) label;
		    sheet.addCell(cell);
	    }

	    WritableSheet sheet_dpa = writableWorkbook.getSheet("DPA_TestSet");
	    int row_dpa = sheet_dpa.getRows();
	    for (int i=1; i<row_dpa; i++) {
		    Label label = new Label(4, i, "");
		    cell = (WritableCell) label;
		    sheet_dpa.addCell(cell);
	    }

	    WritableSheet sheet_qm = writableWorkbook.getSheet("QM_TestSet");
	    int row_qm = sheet_qm.getRows();
	    for (int i=1;i<row_qm;i++) {
		    Label label = new Label(4, i, "");
		    label.setString("");
		    cell = (WritableCell) label;
		    sheet_qm.addCell(cell);
	    }

	    WritableSheet sheet_wfm = writableWorkbook.getSheet("WFM_TestSet");
	    int row_wfm = sheet_wfm.getRows();
	    for (int i=1; i<row_wfm; i++) {
		    Label l = new Label(4, i, "");
		    cell = (WritableCell) l;
		    sheet_wfm.addCell(cell);
	    }

	    writableWorkbook.write();
	    writableWorkbook.close();
    }

	public static void writePassword(WebDriver driver, String pwd, int col, int row) throws Exception {
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
