package winium.setupenvironment;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.winium.DesktopOptions;
import org.openqa.selenium.winium.WiniumDriver;
import org.openqa.selenium.winium.WiniumDriverService;
import outlookaddin.configutrationwindow.ExcelArrayList;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Created by sarja on 5/24/2017.
 */
public class WiniumEnvironmentConfiguration {


    public static WiniumDriver driver = null;
    public static WiniumDriverService service = null;
    public static DesktopOptions options = null;

    int totalNumberOfdataForDirectorySheet;
    int totalNumberOfdataForConnectionPropertySheet;
    int totalNumberOfdataForVariableSheet;
    int totalNumberOfdataForOutlookVersionSheet;

    public List<ExcelOperation> requiredDirectories = new ArrayList<ExcelOperation>();
    public List<ExcelOperation> connectionProperty = new ArrayList<ExcelOperation>();
    public List<ExcelOperation> variablesList = new ArrayList<ExcelOperation>();


    public ExcelArrayList readExcel() throws IOException {


        FileInputStream file = new FileInputStream(new File("WiniumEnvironmentSetupVariable.xlsx"));

//Get the workbook instance for XLS file
        XSSFWorkbook workbook = new XSSFWorkbook(file);

//Get  directory sheet from the workbook
        XSSFSheet directorySheet = workbook.getSheet("Directories");


//Get Inspect element Connection details sheet
        XSSFSheet connectionPropertySheet = workbook.getSheet("Connection details");

//Get outlook version from OutlookVersion sheet
       // XSSFSheet OutlookVersionSheet = workbook.getSheet("OutlookVersion");

//Get variables from variable sheet
        XSSFSheet variableSheet = workbook.getSheet("Variables");



//Get iterator to all the rows in directorySheet
        Iterator<Row> directorySheetRowIterator = directorySheet.iterator();
        totalNumberOfdataForDirectorySheet = directorySheet.getLastRowNum();



//Get iterator to all the rows in Connection details sheet
        Iterator<Row> connectionPropertiesSheetRowIterator = connectionPropertySheet.iterator();
        totalNumberOfdataForConnectionPropertySheet = connectionPropertySheet.getLastRowNum();

//Get iterator to all the rows in variabl;es  sheet
        Iterator<Row> variablesSheetRowIterator = variableSheet.iterator();
        totalNumberOfdataForVariableSheet = variableSheet.getLastRowNum();



//Get iterator to all the rows in OutlookVersion sheet
      /*  Iterator<Row> outlookVersionSheetRowIterator = OutlookVersionSheet.iterator();
        totalNumberOfdataForOutlookVersionSheet = OutlookVersionSheet.getLastRowNum();*/



// Assign data to ExcelOperation object for directory sheet
        ExcelOperation directories;
        while (directorySheetRowIterator.hasNext()) {
            Row directoryRow = directorySheetRowIterator.next();
            if (directoryRow.getRowNum() == 0) {
                continue;
            }

            directories = new ExcelOperation();
            directories.outlookPath = directoryRow.getCell(0).getStringCellValue(); // for outlook.exe path
            directories.winiumDriverPath = directoryRow.getCell(1).getStringCellValue(); // for winium driver path
            requiredDirectories.add(directories);

        }


// Assign data to ExcelOperation object for connection property sheet
        ExcelOperation connectionPropertyOperation;
        while (connectionPropertiesSheetRowIterator.hasNext()) {
            Row connectionPropertyRow = connectionPropertiesSheetRowIterator.next();
            if (connectionPropertyRow.getRowNum() == 0) {
                continue;
            }

            connectionPropertyOperation = new ExcelOperation();
            connectionPropertyOperation.serverName = connectionPropertyRow.getCell(0).getStringCellValue(); // for server name
            connectionPropertyOperation.userName = connectionPropertyRow.getCell(1).getStringCellValue(); // for username
            connectionPropertyOperation.password = connectionPropertyRow.getCell(2).getStringCellValue(); // password
            connectionPropertyOperation.domainName = connectionPropertyRow.getCell(3).getStringCellValue(); // domain name
            connectionPropertyOperation.outlookVersionName = connectionPropertyRow.getCell(4).getStringCellValue(); // outlook version
            connectionProperty.add(connectionPropertyOperation);

        }

// Assign data to ExcelOperation object for variable sheet
        ExcelOperation variblesOperation;
        while (variablesSheetRowIterator.hasNext()) {
            Row variablesRow = variablesSheetRowIterator.next();
            if (variablesRow.getRowNum() == 0) {
                continue;
            }

            variblesOperation = new ExcelOperation();
            variblesOperation.biscomSftRibbonTabName = variablesRow.getCell(2).getStringCellValue(); // for biscom SFT tab name

            variablesList.add(variblesOperation);

        }

/*
        ExcelOperation outlookVersionOperation;
        while (outlookVersionSheetRowIterator.hasNext()) {
            Row outlookVersionRow = outlookVersionSheetRowIterator.next();
            if (outlookVersionRow.getRowNum() == 0) {
                continue;
            }

            outlookVersionOperation = new ExcelOperation();
            outlookVersionOperation.outlookVersionName = outlookVersionRow.getCell(0).getStringCellValue(); // for outlook version
            outlookVersion.add(outlookVersionOperation);

        }*/


        ExcelArrayList excellData = new ExcelArrayList();
        excellData.setConnectionProperty(connectionProperty);
        excellData.setVariableToAutomate(variablesList);
        //excellData.setOutlookVersion(outlookVersion);

        return excellData;
    }


    public WiniumDriver setupEnvironment() throws IOException {

        readExcel();

        String outlookApplicationPath = requiredDirectories.get(0).getOutlookPath();
        String winiumDriverPath = requiredDirectories.get(0).getWiniumDriverPath();

        options = new DesktopOptions(); //Initiate Winium Desktop Options
        options.setApplicationPath(outlookApplicationPath); //Set outlook application path

        File drivePath = new File(winiumDriverPath); //Set winium driver path

        service = new WiniumDriverService.Builder().usingDriverExecutable(drivePath).usingPort(9999).withVerbose(true).withSilent(false).buildDesktopService();
        service.start(); //Build and Start a Winium Driver service
        driver = new WiniumDriver(service, options); //Start a winium driver

        return driver;

    }





}
