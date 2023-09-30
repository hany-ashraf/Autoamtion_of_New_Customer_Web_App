import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import java.io.FileOutputStream;
import java.io.IOException;

public class NewCustomerPage {

    //WebDriverManager.chromedriver().setup();
    //Select select = new Select(WebElement se);
    WebDriver driver;
    Actions actions = new Actions(driver);


    @BeforeClass
    @Parameters({"URL","Browser Type"})
    public void setUp(String url, String browserType){
        //A URL of website
        if(browserType.equalsIgnoreCase("Internet Explorer")){
            driver = new InternetExplorerDriver();
        } else if (browserType.equalsIgnoreCase("Firefox")) {
            driver = new FirefoxDriver();
        } else if (browserType.equalsIgnoreCase("Chrome")) {
            driver = new ChromeDriver();
        }

        driver.get(url);
        driver.manage().window().minimize();
    }
    @Test(priority = 1)
    public void signIn() throws InterruptedException {
        //locator of username
        driver.findElement(By.id("MainContent_txtUsername")).sendKeys("Hany");
        //locator of password
        driver.findElement(By.id("MainContent_txtPassword")).sendKeys("123");
        //Button of login
        driver.findElement(By.id("MainContent_LoginButton")).click();
        //waiting for 2 second
        Thread.sleep(2000);
    }

    /*@Test(priority = 2)
    public void signInVerification(){
        Assert.assertEquals(true, true,"Sign IN is correct !");
        System.out.println("Verify the sign in is correct");
    }*/
    @Test(priority = 3)
    public void selectOrganization() throws InterruptedException {
        //open a dropdown of organizations
        driver.findElement(By.id("MasterUserDis_I")).click();
        Thread.sleep(2000);
        //select your wanted organization
        driver.findElement(By.id("MasterUserDis_DDD_L_LBI7T0")).click();
        Thread.sleep(4000);
    }
    @Test(priority = 4)
    public void openAccount() throws InterruptedException {
        //locator of open page to New Customer
        driver.findElement(By.id("MainContent_OpenAccount")).click();
        Thread.sleep(2000);
    }
    @Test(priority = 5)
    public void customerData() throws InterruptedException {
        //Create an object of action type
        Actions actions = new Actions(driver);
        //Create an object of Select type
        Select customerType = new Select(driver.findElement(By.id("CmbCustomerType")));
        //choose one of list by value of customer type
        customerType.selectByValue("1");
        Thread.sleep(2000);
        //open a dropdown of locations
        driver.findElement(By.id("MainContent_Cmbdepartment_B-1")).click();
        Thread.sleep(1000);
        //locator of which location you wanted
        driver.findElement(By.id("MainContent_Cmbdepartment_DDD_L_LBI0T0")).click();
        //open a dropdown of daily options
        driver.findElement(By.id("MainContent_CmbDaily_B-1")).click();
        Thread.sleep(1000);
        //locator of which daily option you wanted
        driver.findElement(By.id("MainContent_CmbDaily_DDD_L_LBI3T0")).click();
        //set your account number
        driver.findElement(By.cssSelector("input#MainContent_txtAccount_I")).sendKeys("5468");
        //set your local number
        driver.findElement(By.cssSelector("input#MainContent_txtLocal_I")).sendKeys("77");
        Thread.sleep(2000);
        //set you name of customer
        driver.findElement(By.id("MainContent_txtCustomerName_I")).sendKeys("Hany Ashraf");
        Thread.sleep(2000);
        //set customer's address
        driver.findElement(By.id("MainContent_txtcustomeraddress_I")).sendKeys("Nahia");
        Thread.sleep(2000);
        //set customer floor's number
        driver.findElement(By.id("MainContent_txtFloor_I")).sendKeys("10");
        Thread.sleep(2000);
        //locator to choice which meter type you want such as single or three
        Select meterType = new Select(driver.findElement(By.id("CmbMeterType")));
        //select which one
        meterType.selectByValue("1");
        Thread.sleep(2000);
        //Locator to select which version customer wanted
        Select meterVersion = new Select(driver.findElement(By.id("CmbMeterVersions")));
        //set the version
        meterVersion.selectByValue("3");
        Thread.sleep(2000);
        //Locator to choice the customer of activity type
        Select activityType = new Select(driver.findElement(By.id("CmbActivityType")));
        //set which activity type
        activityType.selectByValue("2");
        Thread.sleep(2000);
        //Locator to choice the Location Description of customer
        Select locationDescription = new Select(driver.findElement(By.id("CmbLocationDescription")));
        //set the location Description
        locationDescription.selectByValue("11");
        Thread.sleep(2000);
        //Locator to give availability to choice which transformer customer wanted
        driver.findElement(By.id("MainContent_CmbTransferCode_B-1")).click();
        Thread.sleep(1000);
        //set the desired transformer
        driver.findElement(By.id("MainContent_CmbTransferCode_DDD_L_LBI0T0")).click();
        //set customer initial charge
        driver.findElement(By.id("MainContent_txtinitialValue_I")).sendKeys("100");
        //set customer horse_power
        driver.findElement(By.id("MainContent_txtHoursePower_I")).sendKeys("50");
        Thread.sleep(3000);
        //Locator of a dropdown to select which technician
        driver.findElement(By.id("MainContent_comboBox6_B-1")).click();
        Thread.sleep(2000);
        //set customer technician
        driver.findElement(By.id("MainContent_comboBox6_DDD_L_LBI0T0")).click();
    }
    @Test(priority = 6)
    public void saveData() throws InterruptedException {
        Thread.sleep(2000);
        JavascriptExecutor scroll = (JavascriptExecutor) driver;
        //Scroll up to find save button
        scroll.executeScript("window.scrollBy(0,-50000)");
        Thread.sleep(1000);
        //click on the save button
        driver.findElement(By.id("SaveBtnID")).click();
        Thread.sleep(10000);
        //Click on the close button of pop_up after save
        driver.findElement(By.id("btnClosePrimary")).click();
    }
    @Test(priority = 7)
    public void saveCustomerIDs() throws InterruptedException {
        // Create a new workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        //set name of sheet Excel
        Sheet sheet = workbook.createSheet("New Customers Codes");
        JavascriptExecutor scroll = (JavascriptExecutor) driver;
        //Scroll down to get customer code
        scroll.executeScript("window.scrollBy(0,500)");
        Thread.sleep(5000);
        //Create a webelement to get text of it
        WebElement Id = driver.findElement(By.id("MainContent_txtAccount_I"));
        //save a text in a local variable
        String cellText = Id.getText();
        Thread.sleep(4000);
        //Create row and column
        Row excelRow = sheet.createRow(5);
        Cell cell = excelRow.createCell(5);
        Thread.sleep(2000);
        // Set the cell value as the read-only cell value in Excel sheet
        cell.setCellValue(cellText);
        // Save the workbook to a file
        try (FileOutputStream outputStream = new FileOutputStream("data.xlsx")) {
            workbook.write(outputStream);
        } catch (IOException e) {
            //Input/Output Exception
            e.printStackTrace();
        }finally {
            driver.quit();
        }

    }

    @AfterClass
    public void tearDown() throws Exception{
        Thread.sleep(5000);
        driver.quit();
    }

}
