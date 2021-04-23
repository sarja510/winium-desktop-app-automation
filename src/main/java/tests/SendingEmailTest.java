package tests;

import org.testng.annotations.Parameters;
import org.testng.annotations.Test;
import outlookaddin.sendemail.SendEmail;

import static tests.TestEnvSetup.excelData;
//import static winium.setupenvironment.WiniumEnvironmentConfiguration.driver;
import static  tests.TestEnvSetup.driver;

/**
 * Created by sarja on 5/30/2017.
 */
public class SendingEmailTest {


    @Parameters ({"sendThroughBiscom","recipientMustSignIn","sender","recipient","subject","emailBodyMessage","emailNotificationMessage","passwordProtected","password","attachFileWithSFT" ,"uploadFileLocation", "caseNumber"})
    @Test(description = "Send new secure email directly via SFT", groups = "sendSecureMail", suiteName="Outlook add in general tests")
    public void sendNewSecureEmail(String sendThroughBiscom, String recipientMustSignIn, String sender, String recipient, String subject, String emailBodyMessage, String emailNotificationMessage, String passwordProtected, String password, String attachFileWithSFT, String uploadFileLocation, int caseNumber) throws InterruptedException {

        String secureMail = "Yes";
        String outlookVersion = excelData.connectionProperty.get(0).outlookVersionName;
        String outlookVersionCheckSixteen = "Outlook 2016";
        String outlookVersionCheckThirteen = "Outlook 2013";
        System.out.println(outlookVersion);
        if ((outlookVersion.equals(outlookVersionCheckSixteen)) || (outlookVersion.equals(outlookVersionCheckThirteen))) {

            System.out.println(outlookVersion);
            SendEmail email = new SendEmail();
            email.sendingDeliveryFromNewSecureMail(driver,secureMail,sendThroughBiscom,recipientMustSignIn,sender,recipient,subject,emailBodyMessage,emailNotificationMessage,passwordProtected,password,attachFileWithSFT,uploadFileLocation,caseNumber);

        }
        else {

            SendEmail email = new SendEmail();
            email.sendingDeliveryFromNewSecureMailOutlook2010(driver,sender,recipient,subject,emailBodyMessage,emailNotificationMessage,uploadFileLocation);
        }


    }



    @Parameters ({"sendThroughBiscom","recipientMustSignIn","sender","recipient","subject","emailBodyMessage","emailNotificationMessage","passwordProtected","password","attachFileWithSFT" ,"uploadFileLocation", "caseNumber"})
    @Test(description = "Send normal email with sft rules", groups = "sendSecureMail", suiteName="Outlook add in general tests")
    public void sendNewEmail(String sendThroughBiscom, String recipientMustSignIn, String sender, String recipient, String subject, String emailBodyMessage, String emailNotificationMessage, String passwordProtected, String password, String attachFileWithSFT, String uploadFileLocation, int caseNumber) throws InterruptedException {


        String secureMail = "No";
        String outlookVersion = excelData.connectionProperty.get(0).outlookVersionName;
        String outlookVersionCheckSixteen = "Outlook 2016";
        String outlookVersionCheckThirteen = "Outlook 2013";
        System.out.println(outlookVersion);
        if ((outlookVersion.equals(outlookVersionCheckSixteen)) || (outlookVersion.equals(outlookVersionCheckThirteen))) {

            System.out.println(outlookVersion);
            SendEmail email = new SendEmail();
            email.sendingDeliveryFromNewSecureMail(driver, secureMail, sendThroughBiscom, recipientMustSignIn, sender, recipient, subject, emailBodyMessage, emailNotificationMessage, passwordProtected, password, attachFileWithSFT, uploadFileLocation,caseNumber);

        } else {

            SendEmail email = new SendEmail();
            email.sendingDeliveryFromNewSecureMail(driver, secureMail, sendThroughBiscom, recipientMustSignIn, sender, recipient, subject, emailBodyMessage, emailNotificationMessage, passwordProtected, password, attachFileWithSFT, uploadFileLocation,caseNumber);
        }
    }

}


