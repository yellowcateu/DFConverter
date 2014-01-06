
package components;

import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.*;
import javax.swing.filechooser.*;
/**
 *
 * @author zolty
 */
public class DFConverterMain extends JPanel
                             implements ActionListener {
    static private final String newline = "\n";
    JButton openButton, saveButton;
    JTextArea log;
    JTextArea datafarm;
    JFileChooser fc, fs;

    public DFConverterMain() {
        super(new BorderLayout());

        
        Font font;
        font = new Font("Courier", 2, 10);
        
        log = new JTextArea(5,20);
        log.setMargin(new Insets(5,5,5,5));
        log.setEditable(false);
        log.setFont(font);
        JScrollPane logScrollPane = new JScrollPane(log);

        
        datafarm = new JTextArea(5,20);
        datafarm.setMargin(new Insets(5,5,5,5));
        datafarm.setEditable(false);
        datafarm.setFont(font);
        JScrollPane dfScrollPane = new JScrollPane(datafarm);

        //Create a file chooser
        fc = new JFileChooser();
        fs = new JFileChooser();
        FileNameExtensionFilter filterXLS = new FileNameExtensionFilter("XLS files", "xls");
        //FileNameExtensionFilter filterXLSX = new FileNameExtensionFilter("XLSX files","xlsx");
        fc.setFileFilter(filterXLS);
        //fc.setFileFilter(filterXLSX);

        openButton = new JButton("Otwórz plik XLS",
                                 createImageIcon("images/Open16.gif"));
        openButton.addActionListener(this);

        saveButton = new JButton("Zapisz plik jako DATA-FARM",
                                 createImageIcon("images/Save16.gif"));
        saveButton.addActionListener(this);

        //For layout purposes, put the buttons in a separate panel
        JPanel buttonPanel = new JPanel(); //use FlowLayout
        buttonPanel.add(openButton);
        buttonPanel.add(saveButton);

        //Add the buttons and the log to this panel.
        add(buttonPanel, BorderLayout.PAGE_START);
        add(logScrollPane, BorderLayout.PAGE_END);
        add(dfScrollPane,BorderLayout.CENTER);
    }

    public void actionPerformed(ActionEvent e) {

        
        //Handle open button action.
        if (e.getSource() == openButton) {
            datafarm.setText("");
            int returnVal = fc.showOpenDialog(DFConverterMain.this);

            if (returnVal == JFileChooser.APPROVE_OPTION) {
                File file = fc.getSelectedFile();
                //This is where a real application would open the file.
                log.append("Otwieram: " + file.getName() + "." + newline);
                ReadXLSFile rx = new ReadXLSFile();
                
                try {
                    datafarm.append(rx.StartReadig(file));
                    rx.initVars();
                } catch (FileNotFoundException ex) {
                    Logger.getLogger(DFConverterMain.class.getName()).log(Level.SEVERE, null, ex);
                } catch (IOException ex) {
                    Logger.getLogger(DFConverterMain.class.getName()).log(Level.SEVERE, null, ex);
                }
            } else {
                log.append("Akcja przerwana przez u¿ytkownika" + newline);
            }
            log.setCaretPosition(log.getDocument().getLength());

        //Handle save button action.
        } else if (e.getSource() == saveButton) {
            int returnVal = fs.showSaveDialog(DFConverterMain.this);
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                File file = fs.getSelectedFile();
                log.append("Zapisujê: " + file.getAbsolutePath() + "." + newline);
                    try{

                        FileWriter fstream = new FileWriter(file.getAbsolutePath());
                        BufferedWriter out = new BufferedWriter(fstream);
                        out.write(datafarm.getText());

                        out.close();
                        log.append("Pomyœlnie zapisano plik: "+file.getAbsolutePath()+newline);
                        }catch (Exception exc){
                        log.append("Error: " + exc.getMessage());
                    }
                
            } else {
                log.append("Akcja przerwana przez u¿ytkownika" + newline);
            }
            log.setCaretPosition(log.getDocument().getLength());
        }
    }

    /** Returns an ImageIcon, or null if the path was invalid. */
    protected static ImageIcon createImageIcon(String path) {
        java.net.URL imgURL = DFConverterMain.class.getResource(path);
        if (imgURL != null) {
            return new ImageIcon(imgURL);
        } else {
            System.err.println("Couldn't find file: " + path);
            return null;
        }
    }

    /**
     * Create the GUI and show it.  For thread safety,
     * this method should be invoked from the
     * event dispatch thread.
     */
    private static void createAndShowGUI() {
        //Create and set up the window.
        JFrame frame = new JFrame("DFConverter");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        //Add content to the window.
        frame.add(new DFConverterMain());

        //Display the window.
        frame.pack();
        frame.setVisible(true);
        frame.setExtendedState(Frame.MAXIMIZED_BOTH);  
    }

    public static void main(String[] args) {
        //Schedule a job for the event dispatch thread:
        //creating and showing this application's GUI.
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                //Turn off metal's use of bold fonts
                UIManager.put("swing.boldMetal", Boolean.FALSE); 
                try {
                // Set System L&F
                    UIManager.setLookAndFeel(
                    UIManager.getSystemLookAndFeelClassName());
                    } 
                    catch (UnsupportedLookAndFeelException e) {
                       // handle exception
                    }
                    catch (ClassNotFoundException e) {
                       // handle exception
                    }
                    catch (InstantiationException e) {
                       // handle exception
                    }
                    catch (IllegalAccessException e) {
                       // handle exception
                    }

                createAndShowGUI();
            }
        });
    }
}
