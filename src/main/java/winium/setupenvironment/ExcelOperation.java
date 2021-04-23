package winium.setupenvironment;

/**
 * Created by sarja on 5/24/2017.
 */
public class ExcelOperation {

    public String outlookPath;
    public String winiumDriverPath;

    public String getOutlookPath() {
        return outlookPath;
    }

    public void setOutlookPath(String outlookPath) {
        this.outlookPath = outlookPath;
    }

    public String getWiniumDriverPath() {
        return winiumDriverPath;
    }

    public void setWiniumDriverPath(String winiumDriverPath) {
        this.winiumDriverPath = winiumDriverPath;
    }





    //Some variable for diffrenet outlook

    public String newMailButtonName;

    public String getNewMailButtonName() {
        return newMailButtonName;
    }

    public void setNewMailButtonName(String newMailButtonName) {
        this.newMailButtonName = newMailButtonName;
    }




    //Sft Connection variables

    public String serverName;
    public String userName;
    public String recipientUserName;
    public String password;
    public String domainName;
    public String wrongServerName;
    public String wrongUserName;
    public String wrongRecipientUserName;
    public String wrongPassword;
    public String wrongDomainName;

    public String getWrongServerName() {
        return wrongServerName;
    }

    public void setWrongServerName(String wrongServerName) {
        this.wrongServerName = wrongServerName;
    }

    public String getWrongUserName() {
        return wrongUserName;
    }

    public void setWrongUserName(String wrongUserName) {
        this.wrongUserName = wrongUserName;
    }

    public String getWrongRecipientUserName() {
        return wrongRecipientUserName;
    }

    public void setWrongRecipientUserName(String wrongRecipientUserName) {
        this.wrongRecipientUserName = wrongRecipientUserName;
    }

    public String getWrongPassword() {
        return wrongPassword;
    }

    public void setWrongPassword(String wrongPassword) {
        this.wrongPassword = wrongPassword;
    }

    public String getWrongDomainName() {
        return wrongDomainName;
    }

    public void setWrongDomainName(String wrongDomainName) {
        this.wrongDomainName = wrongDomainName;
    }

    public String getServerName() {
        return serverName;
    }

    public void setServerName(String serverName) {
        this.serverName = serverName;
    }

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public String getPassword() {
        return password;
    }

    public void setPassword(String password) {
        this.password = password;
    }

    public String getDomainName() {
        return domainName;
    }

    public void setDomainName(String domainName) {
        this.domainName = domainName;
    }

    public String getRecipientUserName() {
        return recipientUserName;
    }

    public void setRecipientUserName(String recipientUserName) {
        this.recipientUserName = recipientUserName;
    }


 //set outlookVersionName

    public String outlookVersionName;

    public String getOutlookVersionName() {
        return outlookVersionName;
    }

    public void setOutlookVersionName(String outlookVersionName) {
        this.outlookVersionName = outlookVersionName;
    }


  // automation variables wich frequently change

    public String biscomSftRibbonTabName;


    public String getBiscomSftRibbonTabName() {
        return biscomSftRibbonTabName;
    }

    public void setBiscomSftRibbonTabName(String biscomSftRibbonTabName) {
        this.biscomSftRibbonTabName = biscomSftRibbonTabName;
    }
}
