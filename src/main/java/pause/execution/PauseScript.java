package pause.execution;

import javax.swing.*;

/**
 * Created by sarja on 5/29/2017.
 */
public class PauseScript {

    public static void infoBox(String infoMessage, String titleBar)
    {
        JOptionPane.showMessageDialog(null, infoMessage, "InfoBox: " + titleBar, JOptionPane.INFORMATION_MESSAGE);
    }
}
