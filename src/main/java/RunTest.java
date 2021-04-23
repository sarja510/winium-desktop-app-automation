/**
 * Created by sarja on 5/24/2017.
 */

import org.openqa.selenium.winium.WiniumDriver;
import org.testng.TestNG;
import outlookaddin.configutrationwindow.ExcelArrayList;

import java.util.ArrayList;
import java.util.List;

public class RunTest {


    public static WiniumDriver driver;
    public static ExcelArrayList excelData;
    public int choice;



    public static void main(String[] args) {


    /*        TestListenerAdapter tla = new TestListenerAdapter();
            TestNG testng = new TestNG();
            testng.setTestClasses(new Class[] { SendingEmailTest.class });
            testng.addListener(tla);
            testng.run();
*/

//To run through testng.xnl
        // Create object of TestNG Class
        TestNG runner = new TestNG();

// Create a list of String
        List<String> suitefiles = new ArrayList<String>();

// Add xml file which you have to execute
        suitefiles.add(".\\TestNgXML\\testng.xml");

// now set xml file for execution
        runner.setTestSuites(suitefiles);


// finally execute the runner using run method
        runner.run();


    }






}
