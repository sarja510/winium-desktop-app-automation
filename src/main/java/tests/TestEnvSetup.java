package tests;

import org.openqa.selenium.winium.WiniumDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import outlookaddin.configutrationwindow.ExcelArrayList;
import winium.setupenvironment.WiniumEnvironmentConfiguration;

import java.io.IOException;

/**
 * Created by sarja on 7/13/2017.
 */
public class TestEnvSetup {

    public static WiniumDriver driver = null;
    public static ExcelArrayList excelData;

    @BeforeTest

    public void setUpEnvironment() throws IOException, InterruptedException {


        WiniumEnvironmentConfiguration data = new WiniumEnvironmentConfiguration();
        excelData = data.readExcel();


        WiniumEnvironmentConfiguration setup = new WiniumEnvironmentConfiguration();
        driver = setup.setupEnvironment();
        driver.getWindowHandles();



    }



    @AfterTest
    public void stopDriver() throws InterruptedException {

        try {
            driver.quit();
        } catch (Exception e) {
            e.printStackTrace();
        }
        //driver.close();

    }


}
