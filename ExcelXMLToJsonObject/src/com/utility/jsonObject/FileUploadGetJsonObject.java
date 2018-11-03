package com.utility.jsonObject;

/**
 * @author Mukundan Kannan
 *
 */

import java.awt.BorderLayout;
import java.awt.Insets;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import javax.swing.JButton;
import javax.swing.JComponent;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.UIManager;
import javax.swing.UIManager.LookAndFeelInfo;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.json.JSONObject;
import com.google.gson.JsonObject;


public class FileUploadGetJsonObject extends JPanel implements ActionListener {
    /**
     * 
     */
    private static final long serialVersionUID = 1L;


    JButton openButton, buttonCopy;
    JTextArea resultText;
    static JFrame frame;
    JProgressBar progressBar;
    JFileChooser fc;
    JProgressBar current = new JProgressBar(0, 2000);
    int num = 0;

    public FileUploadGetJsonObject() {
        super(new BorderLayout());

        resultText = new JTextArea(20, 50);
        resultText.setText("Welcome to get the Json Object");
        resultText.setMargin(new Insets(5, 5, 5, 5));
        resultText.setLineWrap(true);
        resultText.setEditable(false);
        JScrollPane logScrollPane = new JScrollPane(resultText);

        // Create a file chooser
        fc = new JFileChooser();

        openButton = new JButton("Select xlsx_or_xml File");
        openButton.addActionListener(this);

        buttonCopy = new JButton("Copy");
        buttonCopy.addActionListener(this);

        // For layout purposes, put the buttons in a separate panel
        JPanel buttonPanel = new JPanel(); // use FlowLayout
        buttonPanel.add(openButton);

        JPanel buttonPanel1 = new JPanel();
        buttonPanel1.add(buttonCopy);

        buttonPanel.add(current);
        current.setValue(0);
        current.setStringPainted(true);

        // Add the buttons and the to this panel.
        add(buttonPanel, BorderLayout.PAGE_START);
        add(logScrollPane, BorderLayout.CENTER);
        add(buttonPanel1, BorderLayout.PAGE_END);


    }

    public void actionPerformed(ActionEvent e) {
        JsonObject jsonObj = null;
        JSONObject xmlTojsonObj = null;
        
        // Handle open button action.
        if (e.getSource() == openButton) {
            int returnVal = fc.showOpenDialog(FileUploadGetJsonObject.this);
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                File file = fc.getSelectedFile();
                String ext1 = FilenameUtils.getExtension(file.getAbsolutePath());
                // To select the respective excel file ext
                if (ext1.equalsIgnoreCase("xls") || ext1.equalsIgnoreCase("xlsx")) {
                    try {
                        jsonObj = CommonUtil.getExcelDataAsJsonObject(file);
                    } catch (InvalidFormatException e1) {
                        e1.printStackTrace();
                    }
                    resultText.setText("");
                    resultText.append("" + CommonUtil.toPrettyFormat(jsonObj.toString()));
                } else if (ext1.equalsIgnoreCase("xml")) {
                    xmlTojsonObj = CommonUtil.getXmlDataAsJsonObject(file);
                    resultText.setText("");
                    resultText.append("" + CommonUtil.toPrettyFormat(xmlTojsonObj.toString()));
                }
                else {
                    resultText.setText("");
                    resultText.append("" + "Invalid File. Please select Excel or Xml");
                }

                this.statusPercent();

            } else {
                resultText.append("Open command cancelled by user.");
            }
            // resultText.setCaretPosition(resultText.getDocument().getLength());
            resultText.setCaretPosition(0);
        }
        if (e.getSource() == buttonCopy) {
            StringSelection stringSelection = new StringSelection(resultText.getText());
            Clipboard clpbrd = Toolkit.getDefaultToolkit().getSystemClipboard();
            clpbrd.setContents(stringSelection, null);
        }
    }


    /**
     * Create the GUI and show it. For thread safety, this method should be
     * invoked from the event-dispatching thread.
     */
    private static void createAndShowGUI() {
        // Nimbus look and feel if not it will display default UI
        try {
            for (LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (Exception e) {
            JFrame.setDefaultLookAndFeelDecorated(true);
            JDialog.setDefaultLookAndFeelDecorated(true);
        }

        // Create and set up the window.
        frame = new JFrame("File_Conversion_To_JsonObject");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        // Create and set up the content pane.
        JComponent newContentPane = new FileUploadGetJsonObject();
        newContentPane.setOpaque(true); // content panes must be opaque
        frame.setContentPane(newContentPane);

        // Display the window.

        frame.pack();
        frame.setVisible(true);
        frame.setLocationRelativeTo(null);

    }

    public void statusPercent() {
        while (num < 2000) {
            current.setValue(num);
            try {
                Thread.sleep(10);
            } catch (InterruptedException e) {
            }
            num += 95;
        }
    }
    public static void main(String[] args) {
        javax.swing.SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                createAndShowGUI();
            }
        });
    }

}
