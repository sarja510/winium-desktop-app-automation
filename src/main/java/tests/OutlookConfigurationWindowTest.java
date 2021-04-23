package tests;

import org.testng.annotations.Parameters;
import org.testng.annotations.Test;
import outlookaddin.configutrationwindow.OutlookConfigurationWindow;
import pause.execution.PauseScript;

import static tests.TestEnvSetup.excelData;
import static winium.setupenvironment.WiniumEnvironmentConfiguration.driver;

/**
 * Created by sarja on 5/29/2017.
 */
public class OutlookConfigurationWindowTest {




    @Test(description = "test", groups = "ConfigurationWindowTest")

    public void emptyCredential() throws InterruptedException {

        String servername = excelData.connectionProperty.get(0).serverName;
        String username = excelData.connectionProperty.get(0).userName;
        String password = excelData.connectionProperty.get(0).password;
        String domain = excelData.connectionProperty.get(0).domainName;


        PauseScript window = new PauseScript();
        window.infoBox("Please enable outlook addin from server and click ok","Pause script");

        OutlookConfigurationWindow emptyConfiguration = new OutlookConfigurationWindow();
        emptyConfiguration.emptyCredentialConnection(driver, servername, username, password, domain);

    }


    @Test(description = "test", groups = "ConfigurationWindowTest")

    public void wrongCredential() throws InterruptedException {


        String servername = excelData.connectionProperty.get(0).serverName;
        String username = excelData.connectionProperty.get(0).userName;
        String password = excelData.connectionProperty.get(0).password;
        String domain = excelData.connectionProperty.get(0).domainName;
        String wrongServername = "1.18.11.5000";
        String wrongUsername = "testest@testtest.com";
        String wrongPassword = "wrongpasswd";
        String wrongDomain = "";

        OutlookConfigurationWindow wrongConfiguration = new OutlookConfigurationWindow();
        wrongConfiguration.wrongCredentialConnection(driver, servername, username, password, domain, wrongServername, wrongUsername, wrongPassword, wrongDomain);


    }



    @Test (description = "test", groups = "ConfigurationWindowTest")
    public void outolookAddinDisabledFromServer() throws InterruptedException{


        String servername = excelData.connectionProperty.get(0).serverName;
        String username = excelData.connectionProperty.get(0).userName;
        String password = excelData.connectionProperty.get(0).password;
        String domain = excelData.connectionProperty.get(0).domainName;

        PauseScript window = new PauseScript();
        window.infoBox("Please disable outlook addin from server and click ok","Pasue script");


        OutlookConfigurationWindow correctConfiguration = new OutlookConfigurationWindow();
        correctConfiguration.correctCredentialConnection(driver, servername, username, password, domain, 2);

    }


    @Parameters({"servername","username","password","domain"})
    @Test (description = "test", groups = "ConfigurationWindowTest")

    public void correctCredential() throws InterruptedException {


        String servername = excelData.connectionProperty.get(0).serverName;
        String username = excelData.connectionProperty.get(0).userName;
        String password = excelData.connectionProperty.get(0).password;
        String domain = excelData.connectionProperty.get(0).domainName;


        PauseScript window = new PauseScript();
        window.infoBox("Please cancel the configuration window and enable outlook addin from server and click ok","Pasue script");

        OutlookConfigurationWindow correctConfiguration = new OutlookConfigurationWindow();
        correctConfiguration.correctCredentialConnection(driver, servername, username, password, domain, 1);


    }




}
