package outlookaddin.sendemail;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.winium.WiniumDriver;

import java.io.File;
import java.io.IOException;

/**
 * Created by sarja on 5/24/2017.
 */
public class SendEmail {



    public void sendingDeliveryFromNewSecureMail(WiniumDriver driver,String secureMail,String SendThroughBiscom, String RecipientMustSignIn, String sender, String recipient, String subject, String emailBodyMessage, String emailNotificationMessage,String passwordProtected, String password, String attachFileWithSFT, String uploadFileLocation, int caseNumber)throws InterruptedException {

        Thread.sleep(4000);
        WebElement outlookFrame = driver.findElement(By.className("rctrl_renwnd32"));
        outlookFrame.findElement(By.name("Home")).click();
        Thread.sleep(2000);


        if (secureMail.equals("Yes")) {

            outlookFrame.findElement(By.name("New Secure Email")).click();
            Thread.sleep(2000);

            //Take screen shot for default email window after clicking New Secure Email
            File scrFileOne = driver.getScreenshotAs(OutputType.FILE);
            try {
                FileUtils.copyFile(scrFileOne, new File(".\\ScreenShots\\SendEmail\\SendEmailThroughNewSecureEmail-DefaultWindow "+caseNumber+".png"));
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        else {

            /*Thread.sleep(10000);*/
            outlookFrame.findElement(By.name("New Email")).click();
            Thread.sleep(2000);

            //Take screen shot for default email window after clicking New Email
            File scrFileOne = driver.getScreenshotAs(OutputType.FILE);
            try {
                FileUtils.copyFile(scrFileOne, new File(".\\ScreenShots\\SendEmail\\SendEmailThroughNewEmail-DefaultWindow "+caseNumber+".png"));
            } catch (IOException e) {
                e.printStackTrace();
            }
        }


// Set Biscom sft rules for email
        if (SendThroughBiscom.equals("Yes")) {

            WebElement biscomSftGroup = driver.findElement(By.name("Biscom SFT"));
            biscomSftGroup.findElement(By.name("Open")).click();
            biscomSftGroup.findElement(By.name("Yes")).click();


        }
        else if (SendThroughBiscom.equals("No")) {

            WebElement biscomSftGroup = driver.findElement(By.name("Biscom SFT"));
            biscomSftGroup.findElement(By.name("Open")).click();
            biscomSftGroup.findElement(By.name("No")).click();

        }
        else if (SendThroughBiscom.equals("Use policy")){

            WebElement biscomSftGroup = driver.findElement(By.name("Biscom SFT"));
            biscomSftGroup.findElement(By.name("Open")).click();
            biscomSftGroup.findElement(By.name("Use policy")).click();
        }

        if (RecipientMustSignIn.equals("Yes")){

            WebElement biscomSftGroup =  driver.findElement(By.xpath("//*[contains(@Name,'Recipients must sign in: ') and contains(@LocalizedControlType,'ComboBox')]"));
            biscomSftGroup.findElement(By.name("Open")).click();
            biscomSftGroup.findElement(By.name("Yes")).click();

        }

        else if (RecipientMustSignIn.equals("No")){


            WebElement biscomSftGroup =  driver.findElement(By.xpath("//*[contains(@Name,'Recipients must sign in: ') and contains(@LocalizedControlType,'ComboBox')]"));
            biscomSftGroup.findElement(By.name("Open")).click();
            biscomSftGroup.findElement(By.name("No")).click();

        }

        Thread.sleep(2000);

        //Take screen shot for email window after setting the rules
        File scrFileOne = driver.getScreenshotAs(OutputType.FILE);
        try {
            FileUtils.copyFile(scrFileOne, new File(".\\ScreenShots\\SendEmail\\EmailSendingRulesSet "+caseNumber+".png"));
        } catch (IOException e) {
            e.printStackTrace();
        }


        WebElement emailFrame = driver.findElement(By.name("Form Regions"));
        WebElement topRibon = driver.findElement(By.name("MsoDockTop"));
        WebElement sftGroup = driver.findElement(By.name("Biscom SFT"));


        //Custom mail address to sent from

        Thread.sleep(2000);
        emailFrame.findElement(By.name("From")).click();

        WebElement menugroup = driver.findElement(By.name("Menu Group"));
        menugroup.findElement(By.name("Other E-mail Address...")).click();


        WebElement otherEmailAddressBox = driver.findElement(By.name("Send From Other E-mail Address"));
        otherEmailAddressBox.findElement(By.name("From")).click();
        otherEmailAddressBox.findElement(By.name("From")).sendKeys(sender);
        otherEmailAddressBox.findElement(By.name("OK")).click();

//Set recipient for email

        emailFrame.findElement(By.name("To")).sendKeys(recipient);
        emailFrame.findElement(By.id("4101")).click();
        emailFrame.findElement(By.id("4101")).sendKeys(subject);
// Set email body message
        WebElement emailFrametwo =  driver.findElement(By.name(subject+" - Message"));
        emailFrametwo.sendKeys(emailBodyMessage);

//Configure other delivery options

        sftGroup.findElement(By.name("Other Options")).click();

        WebElement deliveryOptions = driver.findElement(By.id("DeliveryOptionForm"));

//Set email notification message

        Thread.sleep(3000);

        driver.getWindowHandles();
        deliveryOptions.findElement(By.className("Internet Explorer_Server")).sendKeys(emailNotificationMessage);


        deliveryOptions.findElement(By.name("Show Advanced Options >>")).click();


//Set password for delivery
        if (passwordProtected.equals("Yes")) {

            deliveryOptions.findElement(By.id("txtPassword")).sendKeys(password);
            deliveryOptions.findElement(By.id("txtConfirmPassword")).sendKeys(password);
        }



//Take screen shot for default delivery option window
        File scrFileTwo = driver.getScreenshotAs(OutputType.FILE);
        try {
            FileUtils.copyFile(scrFileTwo, new File(".\\ScreenShots\\SendEmail\\DefaultDeliveryOptionWindow "+caseNumber+".png"));
        } catch (IOException e) {
            e.printStackTrace();
        }

        deliveryOptions.findElement(By.name("OK")).click();
        Thread.sleep(2000);


// Attach file with biscom attach
        if (attachFileWithSFT.equals("Yes")) {
            WebElement addFile = driver.findElement(By.name("Biscom SFT"));
            addFile.findElement(By.name("Attach File")).click();

            WebElement uploadWindow = driver.findElement(By.name("Insert File"));
            uploadWindow.findElement(By.className("Edit")).sendKeys(uploadFileLocation);
            uploadWindow.findElement(By.name("Open")).click();
            Thread.sleep(3000);
        }
//Attach file with normal attach

        else if (attachFileWithSFT.equals("No")) {

            WebElement addFile = driver.findElement(By.name("Include"));
            addFile.findElement(By.name("Attach File")).click();

            WebElement uploadWindow = driver.findElement(By.name("Insert File"));
            uploadWindow.findElement(By.className("Edit")).sendKeys(uploadFileLocation);
            uploadWindow.findElement(By.name("Insert")).click();
            Thread.sleep(3000);

        }

        emailFrame.findElement(By.name("Send")).click();

        Thread.sleep(10000);

        

    }

    public void sendingDeliveryFromNewSecureMailOutlook2010(WiniumDriver driver, String sender, String recipient, String subject, String emailBodyMessage, String emailNotificationMessage, String uploadFileLocation)throws InterruptedException {

        Thread.sleep(2000);
        WebElement outlookFrame = driver.findElement(By.className("rctrl_renwnd32"));
        outlookFrame.findElement(By.name("Home")).click();
        outlookFrame.findElement(By.name("New Secure E-mail")).click();



        File scrFileOne = driver.getScreenshotAs(OutputType.FILE);
        try {
            FileUtils.copyFile(scrFileOne, new File(".\\ScreenShots\\SendEmail\\SendEmailThroughNewSecureEmail-1.png"));
        } catch (IOException e) {
            e.printStackTrace();
        }


        WebElement emailFrmae = driver.findElement(By.name("Form Regions"));
        WebElement topRibon = driver.findElement(By.name("MsoDockTop"));
        WebElement sftGroup = driver.findElement(By.name("Biscom SFT"));


        //Custom mail address to sent from

        emailFrmae.findElement(By.name("From")).click();

        WebElement menugroup = driver.findElement(By.className("NetUIListView"));
        menugroup.findElement(By.name("Other E-mail Address...")).click();


        WebElement otherEmailAddressBox = driver.findElement(By.name("Send From Other E-mail Address"));
        otherEmailAddressBox.findElement(By.name("From")).click();
        otherEmailAddressBox.findElement(By.name("From")).sendKeys(sender);
        otherEmailAddressBox.findElement(By.name("OK")).click();

        //topRibon.findElement(By.name("Maximize")).click();

        emailFrmae.findElement(By.name("To")).sendKeys(recipient);
        emailFrmae.findElement(By.id("4101")).click();
        emailFrmae.findElement(By.id("4101")).sendKeys(subject);

        WebElement emailMessageFrmae = driver.findElement(By.name("Message"));

        emailFrmae.findElement(By.className("_WwG")).click();
        emailFrmae.findElement(By.className("_WwG")).sendKeys(emailBodyMessage);

        sftGroup.findElement(By.name("Other Options")).click();


        WebElement deliveryOptions = driver.findElement(By.name("Delivery Options"));


        deliveryOptions.findElement(By.id("txtBody2")).click();
        deliveryOptions.findElement(By.id("txtBody2")).sendKeys(emailNotificationMessage);

        deliveryOptions.findElement(By.name("Show Advanced Options >>")).click();


        File scrFileTwo = driver.getScreenshotAs(OutputType.FILE);
        try {
            FileUtils.copyFile(scrFileTwo, new File(".\\ScreenShots\\SendEmail\\DefaultDeliveryOptionWindow.png"));
        } catch (IOException e) {
            e.printStackTrace();
        }

        deliveryOptions.findElement(By.name("OK")).click();
        Thread.sleep(2000);


        // attach fiel with biscom attach
        WebElement addFile = driver.findElement(By.name("Include"));
        addFile.findElement(By.name("Attach File")).click();

        WebElement uploadWindow = driver.findElement(By.name("Insert File"));
        uploadWindow.findElement(By.className("Edit")).sendKeys(uploadFileLocation);
        uploadWindow.findElement(By.name("Insert")).click();
        Thread.sleep(3000);



        emailFrmae.findElement(By.name("Send")).click();
        Thread.sleep(2000);



    }

    public void sendingDeliveryFromNewMail(WiniumDriver driver,String SendThroughBiscom, String RecipientMustSignIn, String sender, String recipient, String subject, String emailBodyMessage, String emailNotificationMessage,String passwordProtected, String password, String attachFileWithSFT, String uploadFileLocation)throws InterruptedException {

        Thread.sleep(2000);
        WebElement outlookFrame = driver.findElement(By.className("rctrl_renwnd32"));
        outlookFrame.findElement(By.name("Home")).click();
        Thread.sleep(2000);
        outlookFrame.findElement(By.name("New Email")).click();
        Thread.sleep(2000);


        //Take screen shot for default email window after clicking New Secure Email
        File scrFileOne = driver.getScreenshotAs(OutputType.FILE);
        try {
            FileUtils.copyFile(scrFileOne, new File(".\\ScreenShots\\SendEmail\\SendEmailThroughNewEmail-DefaultWindow.png"));
        } catch (IOException e) {
            e.printStackTrace();
        }

// Set Biscom sft rules for email
        if (SendThroughBiscom.equals("Yes")) {

            WebElement biscomSftGroup = driver.findElement(By.name("Biscom SFT"));
            biscomSftGroup.findElement(By.name("Open")).click();
            biscomSftGroup.findElement(By.name("Yes")).click();


        }
        else if (SendThroughBiscom.equals("No")) {

            WebElement biscomSftGroup = driver.findElement(By.name("Biscom SFT"));
            biscomSftGroup.findElement(By.name("Open")).click();
            biscomSftGroup.findElement(By.name("No")).click();

        }
        else if (SendThroughBiscom.equals("Use policy")){

            WebElement biscomSftGroup = driver.findElement(By.name("Biscom SFT"));
            biscomSftGroup.findElement(By.name("Open")).click();
            biscomSftGroup.findElement(By.name("Use policy")).click();
        }

        if (RecipientMustSignIn.equals("Yes")){

            WebElement biscomSftGroup =  driver.findElement(By.xpath("//*[contains(@Name,'Recipients must sign in: ') and contains(@LocalizedControlType,'ComboBox')]"));
            biscomSftGroup.findElement(By.name("Open")).click();
            biscomSftGroup.findElement(By.name("Yes")).click();

        }

        else if (RecipientMustSignIn.equals("No")){


            WebElement biscomSftGroup =  driver.findElement(By.xpath("//*[contains(@Name,'Recipients must sign in: ') and contains(@LocalizedControlType,'ComboBox')]"));
            biscomSftGroup.findElement(By.name("Open")).click();
            biscomSftGroup.findElement(By.name("No")).click();

        }


        WebElement emailFrame = driver.findElement(By.name("Form Regions"));
        WebElement topRibon = driver.findElement(By.name("MsoDockTop"));
        WebElement sftGroup = driver.findElement(By.name("Biscom SFT"));


        //Custom mail address to sent from

        Thread.sleep(2000);
        emailFrame.findElement(By.name("From")).click();

        WebElement menugroup = driver.findElement(By.name("Menu Group"));
        menugroup.findElement(By.name("Other E-mail Address...")).click();


        WebElement otherEmailAddressBox = driver.findElement(By.name("Send From Other E-mail Address"));
        otherEmailAddressBox.findElement(By.name("From")).click();
        otherEmailAddressBox.findElement(By.name("From")).sendKeys(sender);
        otherEmailAddressBox.findElement(By.name("OK")).click();

//Set recipient for email

        emailFrame.findElement(By.name("To")).sendKeys(recipient);
        emailFrame.findElement(By.id("4101")).click();
        emailFrame.findElement(By.id("4101")).sendKeys(subject);
// Set email body message
        WebElement emailFrametwo =  driver.findElement(By.name(subject+" - Message"));
        emailFrametwo.sendKeys(emailBodyMessage);

//Configure other delivery options

        sftGroup.findElement(By.name("Other Options")).click();

        WebElement deliveryOptions = driver.findElement(By.id("DeliveryOptionForm"));

//Set email notification message

        deliveryOptions.findElement(By.className("Internet Explorer_Server")).sendKeys(emailNotificationMessage);

        deliveryOptions.findElement(By.name("Show Advanced Options >>")).click();


//Set password for delivery
        if (passwordProtected.equals("Yes")) {

            deliveryOptions.findElement(By.id("txtPassword")).sendKeys(password);
            deliveryOptions.findElement(By.id("txtConfirmPassword")).sendKeys(password);
        }


//Take screen shot for default delivery option window
        File scrFileTwo = driver.getScreenshotAs(OutputType.FILE);
        try {
            FileUtils.copyFile(scrFileTwo, new File(".\\ScreenShots\\SendEmail\\DefaultDeliveryOptionWindow.png"));
        } catch (IOException e) {
            e.printStackTrace();
        }

        deliveryOptions.findElement(By.name("OK")).click();
        Thread.sleep(2000);

// Attach file with biscom attach
        if (attachFileWithSFT.equals("Yes")) {
            WebElement addFile = driver.findElement(By.name("Biscom SFT"));
            addFile.findElement(By.name("Attach File")).click();

            WebElement uploadWindow = driver.findElement(By.name("Insert File"));
            uploadWindow.findElement(By.className("Edit")).sendKeys(uploadFileLocation);
            uploadWindow.findElement(By.name("Open")).click();
            Thread.sleep(3000);
        }
//Attach file with normal attach

        else if (attachFileWithSFT.equals("No")) {

            WebElement addFile = driver.findElement(By.name("Include"));
            addFile.findElement(By.name("Attach File")).click();

            WebElement uploadWindow = driver.findElement(By.name("Insert File"));
            uploadWindow.findElement(By.className("Edit")).sendKeys(uploadFileLocation);
            uploadWindow.findElement(By.name("Insert")).click();
            Thread.sleep(3000);

        }

        emailFrame.findElement(By.name("Send")).click();
        Thread.sleep(2000);




    }





}
