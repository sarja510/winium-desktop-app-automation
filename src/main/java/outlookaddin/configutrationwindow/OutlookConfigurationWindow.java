package outlookaddin.configutrationwindow;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.winium.WiniumDriver;

import java.io.File;
import java.io.IOException;

import static tests.TestEnvSetup.excelData;


/**
 * Created by sarja on 5/23/2017.
 */
public class OutlookConfigurationWindow {


    String biscomSFTTabName;

    public OutlookConfigurationWindow(){

        biscomSFTTabName = excelData.variableToAutomate.get(0).biscomSftRibbonTabName;

    }


    public void emptyCredentialConnection(WiniumDriver driver, String servername, String username, String password, String domain) throws InterruptedException {

        //System.out.println(a);
        //String biscomSFTTabName = excelData.variableToAutomate.get(0).biscomSftRibbonTabName;
        Thread.sleep(2000);
        WebElement outlookFrame = driver.findElement(By.className("rctrl_renwnd32"));
        outlookFrame.findElement(By.name(biscomSFTTabName)).click();
        outlookFrame.findElement(By.name("Configuration")).click();
        WebElement sftConfiguration = driver.findElement(By.name("Biscom SFT Configuration"));

        sftConfiguration.findElement(By.id("authenticationModeDropdown")).click();
        sftConfiguration.findElement(By.name("Biscom SFT Authentication")).click();
        sftConfiguration.findElement(By.id("txtServer")).clear();
        sftConfiguration.findElement(By.id("txtUsername")).clear();
        sftConfiguration.findElement(By.id("txtPassword")).clear();
        sftConfiguration.findElement(By.id("btnApply")).click();

        WebElement popupFrame = driver.findElement(By.name("Invalid Input"));

        File scrFileOne = driver.getScreenshotAs(OutputType.FILE);
        try {
            FileUtils.copyFile(scrFileOne, new File(".\\ScreenShots\\OutolookConfigurationWindow\\EmptyCredentialConnection1.png"), true);
        } catch (IOException e) {
            e.printStackTrace();
        }
        popupFrame.findElement(By.name("OK")).click();

        sftConfiguration.findElement(By.id("txtServer")).sendKeys(servername);
        sftConfiguration.findElement(By.id("txtUsername")).sendKeys(username);
        sftConfiguration.findElement(By.id("txtPassword")).clear();
        sftConfiguration.findElement(By.id("btnApply")).click();


        WebElement popupFrameTwo = driver.findElement(By.name("Invalid Input"));

        File scrFileTwo = driver.getScreenshotAs(OutputType.FILE);
        try {
            FileUtils.copyFile(scrFileTwo, new File(".\\ScreenShots\\OutolookConfigurationWindow\\EmptyCredentialConnection2.png"));
        } catch (IOException e) {
            e.printStackTrace();
        }

        popupFrameTwo.findElement(By.name("OK")).click();

        sftConfiguration.findElement(By.id("txtUsername")).sendKeys(username);
        sftConfiguration.findElement(By.id("txtServer")).clear();
        sftConfiguration.findElement(By.id("txtPassword")).clear();
        sftConfiguration.findElement(By.id("btnApply")).click();

        WebElement popupFrameThree = driver.findElement(By.name("Invalid Input"));

        File scrFileThree = driver.getScreenshotAs(OutputType.FILE);
        try {
            FileUtils.copyFile(scrFileThree, new File(".\\ScreenShots\\OutolookConfigurationWindow\\EmptyCredentialConnection3.png"));
        } catch (IOException e) {
            e.printStackTrace();
        }


        popupFrameThree.findElement(By.name("OK")).click();

        sftConfiguration.findElement(By.name("Cancel")).click();


        Thread.sleep(3000);
    }

    public void wrongCredentialConnection(WiniumDriver driver, String servername, String username, String password, String domain, String wrongServername, String wrongUsername, String wrongPassword, String wrongDomainName) throws InterruptedException {

        Thread.sleep(2000);
        WebElement outlookFrame = driver.findElement(By.className("rctrl_renwnd32"));
        outlookFrame.findElement(By.name("BISCOM SFT")).click();
        outlookFrame.findElement(By.name("Configuration")).click();
        WebElement sftConfiguration = driver.findElement(By.name("Biscom SFT Configuration"));

        sftConfiguration.findElement(By.id("authenticationModeDropdown")).click();
        sftConfiguration.findElement(By.name("Biscom SFT Authentication")).click();
        sftConfiguration.findElement(By.id("txtServer")).sendKeys(wrongServername);
        sftConfiguration.findElement(By.id("txtUsername")).sendKeys(username);
        sftConfiguration.findElement(By.id("txtPassword")).sendKeys(password);
        sftConfiguration.findElement(By.id("btnApply")).click();


        try {


            WebElement popupFrame = driver.findElement(By.name("Working ..."));
            String name = popupFrame.getAttribute("Name");
            do {

                System.out.print(popupFrame.getAttribute("Name"));

            } while (name == "Working ...");
        } catch (NoSuchElementException e) {
            e.printStackTrace();
        }

        Thread.sleep(2000);
        File scrFile = driver.getScreenshotAs(OutputType.FILE);
        try {
            FileUtils.copyFile(scrFile, new File(".\\ScreenShots\\OutolookConfigurationWindow\\WrongCredentialConnection1.png"));
        } catch (IOException e) {
            e.printStackTrace();
        }


        Thread.sleep(2000);
        sftConfiguration.findElement(By.id("txtServer")).sendKeys(servername);
        sftConfiguration.findElement(By.id("txtUsername")).sendKeys(wrongUsername);
        sftConfiguration.findElement(By.id("txtPassword")).sendKeys(password);
        sftConfiguration.findElement(By.id("btnApply")).click();

        try {
            WebElement popupFrame = driver.findElement(By.name("Working ..."));
            String name = popupFrame.getAttribute("Name");
            do {

                System.out.print(popupFrame.getAttribute("Name"));

            } while (name == "Working ...");

        } catch (NoSuchElementException e) {
            e.printStackTrace();
        }


        Thread.sleep(2000);
        scrFile = driver.getScreenshotAs(OutputType.FILE);
        try {
            FileUtils.copyFile(scrFile, new File(".\\ScreenShots\\OutolookConfigurationWindow\\WrongCredentialConnection2.png"));
        } catch (IOException e) {
            e.printStackTrace();
        }

        Thread.sleep(2000);
        sftConfiguration.findElement(By.id("txtServer")).sendKeys(servername);
        sftConfiguration.findElement(By.id("txtUsername")).sendKeys(username);
        sftConfiguration.findElement(By.id("txtPassword")).sendKeys(wrongPassword);
        sftConfiguration.findElement(By.id("btnApply")).click();


        try {
            WebElement popupFrame = driver.findElement(By.name("Working ..."));
            String name = popupFrame.getAttribute("Name");
            do {

                System.out.print(popupFrame.getAttribute("Name"));

            } while (name == "Working ...");

        } catch (NoSuchElementException e) {
            e.printStackTrace();
        }
        Thread.sleep(2000);
        scrFile = driver.getScreenshotAs(OutputType.FILE);
        try {
            FileUtils.copyFile(scrFile, new File(".\\ScreenShots\\OutolookConfigurationWindow\\WrongCredentialConnection3.png"));
        } catch (IOException e) {
            e.printStackTrace();
        }

        Thread.sleep(2000);
        sftConfiguration.findElement(By.name("Cancel")).click();


    }

    public void correctCredentialConnection(WiniumDriver driver, String servername, String username, String password, String domain, int i) throws InterruptedException {


        Thread.sleep(2000);
        WebElement outlookFrame = driver.findElement(By.className("rctrl_renwnd32"));
        outlookFrame.findElement(By.name("BISCOM SFT")).click();
        outlookFrame.findElement(By.name("Configuration")).click();
        WebElement sftConfiguration = driver.findElement(By.name("Biscom SFT Configuration"));

        sftConfiguration.findElement(By.id("authenticationModeDropdown")).click();
        sftConfiguration.findElement(By.name("Biscom SFT Authentication")).click();
        sftConfiguration.findElement(By.id("txtServer")).sendKeys(servername);
        sftConfiguration.findElement(By.id("txtUsername")).sendKeys(username);
        sftConfiguration.findElement(By.id("txtPassword")).sendKeys(password);
        sftConfiguration.findElement(By.id("btnApply")).click();

        WebElement popupFrame = driver.findElement(By.name("Working ..."));
        String name = popupFrame.getAttribute("Name");
        do {

            System.out.print(popupFrame.getAttribute("Name"));

        } while (name == "Working ...");

        Thread.sleep(2000);
        File scrFileOne = driver.getScreenshotAs(OutputType.FILE);
        try {
            FileUtils.copyFile(scrFileOne, new File(".\\ScreenShots\\OutolookConfigurationWindow\\CorrectCredentialConnection"+i+".png"));
        } catch (IOException e) {
            e.printStackTrace();
        }

        Thread.sleep(2000);
        sftConfiguration.findElement(By.name("OK")).click();

        Thread.sleep(4000);

    }
}
