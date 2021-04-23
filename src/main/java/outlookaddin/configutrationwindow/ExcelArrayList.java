package outlookaddin.configutrationwindow;

import winium.setupenvironment.ExcelOperation;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by sarja on 5/24/2017.
 */
public class ExcelArrayList {

    public List<ExcelOperation> connectionProperty = new ArrayList<ExcelOperation>();

    public  List<ExcelOperation> outlookVersion = new ArrayList<ExcelOperation>();

    public List<ExcelOperation> getOutlookVersion() {
        return outlookVersion;
    }

    public void setOutlookVersion(List<ExcelOperation> outlookVersion) {
        this.outlookVersion = outlookVersion;
    }

    public List<ExcelOperation> getConnectionProperty() {
        return connectionProperty;
    }

    public void setConnectionProperty(List<ExcelOperation> connectionProperty) {
        this.connectionProperty = connectionProperty;
    }

    public void setVariableToAutomate(List<ExcelOperation> variableToAutomate) {
        this.variableToAutomate = variableToAutomate;
    }

    public List<ExcelOperation> variableToAutomate = new ArrayList<ExcelOperation>();
}
