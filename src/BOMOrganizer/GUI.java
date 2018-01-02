package BOMOrganizer;


import org.apache.poi.hssf.model.*;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.*;
import java.util.ArrayList;

import java.util.Random;

/**
 * Created by rahul on 2017-08-15.
 */

public class GUI extends JFrame {

    //class variables

    private JTextField fileOneField; //name of file to be read
    private JTextField fileTwoField; //name of file to be created
    private JTextField fileTwoDirField;

    private JPanel panel;
    private JButton helpButton;
    private JTextArea area;

    private JTextField fileOneDesField; //cell address of first Designator
    private JTextField fileOneComField; //cell address of first Part Number/Comment
    private JTextField fileOneQtyField; //cell address of first Quantity

    private JButton startButton; //start button

    String fileOneString = ""; //name of file to be edited, used by file chooser gui
    File directory; //directory of file, used by file chooser gui
    File directoryTwo;

    GUI() { //constructor
        super("BOMOrganizer"); //title
        this.setSize(1000, 400); //default size
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE); //default close
        this.setResizable(true); //resizable

        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));

        JLabel instructionsLabel = new JLabel("Select a BOM File and Enter New File Name");

        panel.add(instructionsLabel);

        JPanel inputPanel = new JPanel();
        inputPanel.setLayout(new BoxLayout(inputPanel, BoxLayout.Y_AXIS));

        JPanel fileOnePanel = new JPanel(); //panel to enter file one name
        fileOnePanel.setLayout(new FlowLayout());
        JLabel fileOneLabel = new JLabel("BOM File (Excel)");
        fileOneField = new JTextField(40);
        fileOnePanel.add(fileOneLabel);
        fileOnePanel.add(fileOneField);
        JButton fileOneBut = new JButton("Select"); //button to open file chooser gui
        fileOneBut.addActionListener(new FileOneButListener());
        fileOnePanel.add(fileOneBut);

        JPanel fileOneInputPanel = new JPanel(); //panel to get starting cells
        fileOneInputPanel.setLayout(new FlowLayout());

        JLabel fileOneDesStart = new JLabel("Starting Component ID Cell:");
        fileOneDesField = new JTextField(5);
        fileOneInputPanel.add(fileOneDesStart);
        fileOneInputPanel.add(fileOneDesField);

        JLabel fileOneComStart = new JLabel("Starting Part Number Cell:");
        fileOneComField = new JTextField(5);
        fileOneInputPanel.add(fileOneComStart);
        fileOneInputPanel.add(fileOneComField);

        JLabel fileOneQtyStart = new JLabel("Starting Quantity Cell:");
        fileOneQtyField = new JTextField(5);
        fileOneInputPanel.add(fileOneQtyStart);
        fileOneInputPanel.add(fileOneQtyField);

        JPanel fileTwoPanel = new JPanel(); //get new file name
        fileTwoPanel.setLayout(new FlowLayout());
        JLabel fileTwoLabel = new JLabel("Modified File Name (Excel)");
        fileTwoField = new JTextField(40);
        JButton fileTwoBut = new JButton("Select");
        fileTwoBut.addActionListener(new FileTwoButtonListener());
        fileTwoDirField = new JTextField(40);
        JLabel fileTwoDirLabel = new JLabel("Modified File Directory");

        JPanel botPanel = new JPanel();
        botPanel.setLayout(new FlowLayout());

        botPanel.add(fileTwoLabel);
        botPanel.add(fileTwoField);

        fileTwoPanel.add(fileTwoDirLabel);
        fileTwoPanel.add(fileTwoDirField);
        fileTwoPanel.add(fileTwoBut);

        //adds all pannels
        inputPanel.add(fileOnePanel);
        inputPanel.add(fileOneInputPanel);
        inputPanel.add(fileTwoPanel);
        inputPanel.add(botPanel);


        panel.add(inputPanel); //adds to main panel


        JPanel startPanel = new JPanel(new FlowLayout());
        startButton = new JButton("Start"); //to start program
        startButton.addActionListener(new StartButtonListener());
        startPanel.add(startButton);

        helpButton = new JButton("Help");
        helpButton.addActionListener(new HelpButtonListener());
        startPanel.add(helpButton);

        panel.add(startPanel);

        this.add(panel); //adds everything to the gui
        this.setVisible(true);//makes visible
    }


    public class StartButtonListener implements ActionListener {

        String fieldOne; //file name
        String fieldTwo; //new file name
        //starting cells
        String desStartCell;
        String comStartCell;
        String qtyStartCell;


        @Override
        public void actionPerformed(ActionEvent e) {
            //get info from fields
            fieldOne = fileOneField.getText();
            fieldTwo = fileTwoField.getText();
            if (fieldTwo.indexOf(".") != -1 && fieldTwo.indexOf(".xlsx") == -1 && fieldTwo.indexOf(".xls") == -1){
                fieldTwo = "";
            }
            if (fieldTwo.indexOf(".xlsx") == -1 || fieldTwo.indexOf(".xls") == -1){
                fieldTwo += ".xls";
            }
            desStartCell = fileOneDesField.getText().toUpperCase();
            comStartCell = fileOneComField.getText().toUpperCase();
            qtyStartCell = fileOneQtyField.getText().toUpperCase();


            //error check
            if (fieldOne.length() > 4 && fieldTwo.length() > 5) { //make sure not too short ".xlsx" is 5 characters, so name must be longer than that
                if (fieldOne.substring(fieldOne.length() - 5).equals(".xlsx") || (fieldOne.substring(fieldOne.length()-4).equalsIgnoreCase(".xls"))) { //makes sure the end of file name has ".xlsx"
                    try {
                        //make file
                        File file = new File(fieldOne); //check if fieldOne and fieldTwo work instead later
                        InputStream fs = new FileInputStream(file); //input stream
                        Workbook wb = WorkbookFactory.create(fs);
                        Sheet sheet = wb.getSheetAt(0);

                        //makes variables
                        Row row;
                        Cell cell;
                        Cell startingDesCell = null;
                        Cell startingComCell = null;
                        Cell startingQtyCell = null;
                        Cell desCell = null;
                        Cell comCell = null;

                        int rows; // No of rows
                        rows = sheet.getPhysicalNumberOfRows();

                        int cols = 0; // No of columns
                        int tmp = 0; //used to get columns

                        //used to find starting cell
                        String desStartNum = "";
                        String comStartNum = "";
                        String qtyStartNum = "";
                        String desStartLet = "";
                        String comStartLet = "";
                        String qtyStartLet = "";

                        ArrayList<Part> parts = new ArrayList<Part>();//list of unique parts to group designators

                        //splits cell addresses in parts
                        desStartLet = getLetters(desStartCell);
                        desStartNum = getNumbers(desStartCell);

                        comStartLet = getLetters(comStartCell);
                        comStartNum = getNumbers(comStartCell);

                        qtyStartLet = getLetters(qtyStartCell);
                        qtyStartNum = getNumbers(qtyStartCell);


                        //checks if the addresses are on the same row and that none of them are in the same column
                        if (desStartNum.equalsIgnoreCase(comStartNum) && desStartNum.equalsIgnoreCase(qtyStartNum) && !desStartLet.equalsIgnoreCase(qtyStartLet) && !desStartLet.equalsIgnoreCase(comStartLet)) {


                            //used to count the number of columns
                            for (int i = 0; i < 10 || i < rows; i++) {
                                row = sheet.getRow(i);
                                if (row != null) {
                                    tmp = sheet.getRow(i).getPhysicalNumberOfCells();
                                    if (tmp > cols) {
                                        cols = tmp;
                                    }
                                }
                            }

                            int loopCount = 0; //used to make loop go 3 times
                            for (int r = 0; r < rows && loopCount < 3; r++) { //when all 3 are found, loop will stop
                                row = sheet.getRow(r); //gets row
                                if (row != null) {
                                    for (int c = 0; c < cols; c++) {
                                        cell = row.getCell(c);
                                        if (cell != null) {
                                            CellRangeAddress range = new CellRangeAddress(cell.getRowIndex(), cell.getRowIndex(), cell.getColumnIndex(), cell.getColumnIndex());
                                            String rangeString = range.toString(); //gets address of cell
                                            if (rangeString.indexOf(desStartCell) != -1) { //starting designator cell
                                                startingDesCell = cell;
                                                loopCount++;
                                            } else if (rangeString.indexOf(comStartCell) != -1) { //starting part number cell
                                                startingComCell = cell;
                                                loopCount++;
                                            } else if (rangeString.indexOf(qtyStartCell) != -1) { //start quantity cell
                                                startingQtyCell = cell;
                                                loopCount++;
                                            }
                                        }
                                    }
                                }
                            }

                            if (startingDesCell != null && startingComCell != null) {
                                ArrayList<String> originalPartNumbers = new ArrayList<String>(); //list of unique part numbers, so duplicates can be found
                                for (int r = startingDesCell.getRowIndex(); r < rows; r++) { //starts from row where user entered
                                    System.out.println(r);

                                    row = sheet.getRow(r);
                                    if (row != null) {
                                        //gets cells
                                        desCell = row.getCell(startingDesCell.getColumnIndex());
                                        comCell = row.getCell(startingComCell.getColumnIndex());

                                        //sets initial values
                                        String desCellCon = "";
                                        String comCellCon = "";

                                        if (desCell != null && comCell != null) {
                                            //gets values
                                            desCell.setCellType(Cell.CELL_TYPE_STRING);
                                            desCellCon = desCell.getStringCellValue();
                                            comCell.setCellType(Cell.CELL_TYPE_STRING);
                                            comCellCon = comCell.getStringCellValue();


                                            String currentPartNumber = comCellCon; //used to go through excel file

                                            boolean found = false; //value to see if the part number has been found yet
                                            for (int i = 0; i < originalPartNumbers.size(); i++) {
                                                if (originalPartNumbers.get(i).equalsIgnoreCase(currentPartNumber)) { //searches for part number from unique list
                                                    found = true; //found, so goes to next row
                                                }
                                            }
                                            if (!found) { //if not found
                                                if (currentPartNumber.length() > 2) { //avoids errors
                                                    originalPartNumbers.add(currentPartNumber); //adds part number to list of unique ones
                                                }
                                                ArrayList<String> designators = new ArrayList<String>(); //list of designators with same part number
                                                ArrayList<String> otherInfo = new ArrayList<String>(); //other info that comes with the row of a part, varies with different files so everything is grouped
                                                designators.add(desCellCon); //adds initial designator
                                                for (int rowNum = comCell.getRowIndex(); rowNum < rows; rowNum++) { //starts from current row that was found
                                                    Row newRow = sheet.getRow(rowNum); //used to compare
                                                    //gets new cells
                                                    Cell newDesCell = newRow.getCell(startingDesCell.getColumnIndex());
                                                    Cell newComCell = newRow.getCell(startingComCell.getColumnIndex());

                                                    if (newDesCell != null && newComCell != null) {
                                                        //sets values
                                                        newDesCell.setCellType(Cell.CELL_TYPE_STRING);
                                                        String newDesCellCon = newDesCell.getStringCellValue();
                                                        newComCell.setCellType(Cell.CELL_TYPE_STRING);
                                                        String newComCellCon = newComCell.getStringCellValue();


                                                        if (currentPartNumber.equalsIgnoreCase(newComCellCon) && currentPartNumber.length() > 0 && !newDesCellCon.equalsIgnoreCase(desCellCon)) { //error check
                                                            designators.add(newDesCellCon); //adds designator to list
                                                        }
                                                    }

                                                    //gets all other info
                                                    for (int c = 0; c < cols; c++) {
                                                        Cell tempCell = newRow.getCell(c);
                                                        if (tempCell != null) {
                                                            tempCell.setCellType(Cell.CELL_TYPE_STRING);
                                                            String tempCellCon = tempCell.getStringCellValue();
                                                            otherInfo.add(tempCellCon);
                                                        } else {
                                                            otherInfo.add(""); //if there is no other info
                                                        }
                                                    }
                                                }
                                                if (designators.get(0).length() > 0 && currentPartNumber != null && otherInfo != null) { //error check
                                                    Part newPart = new Part(designators, currentPartNumber, otherInfo); //creates part object. Inlcudes list of designators, part number, and other info
                                                    parts.add(newPart); //adds to list of unique parts
                                                }
                                            }
                                        }
                                    }
                                }
                            } else {
                                JOptionPane.showMessageDialog(null, "Cell Entered is Invald");
                            }
                        } else {
                            JOptionPane.showMessageDialog(null, "Cells are not in same row");
                        }

                        /***Writing to new excel file***/
                        try {
                            XSSFWorkbook newBook = new XSSFWorkbook(); //new work book
                            XSSFSheet newSheet = newBook.createSheet("new sheet"); //creates new sheet

                            int numNewRows = parts.size() + startingComCell.getRowIndex(); //calculates new number of rows, should be less than original

                            ArrayList<Part> blanks = new ArrayList<Part>(); //list of parts with no part number
                            for (int i = 0; i < parts.size(); i++) { //goes through list of unique parts
                                if (parts.get(i).getPartNumber().length() < 1) { //finds blank part numbers
                                    blanks.add(parts.get(i)); //adds to list of blank part numbers
                                    parts.remove(i); //removes from part list
                                }
                            }
                            for (int i = 0; i < blanks.size(); i++) {
                                parts.add(blanks.get(i)); //re-adds the blank part numbers so that they appear at the bottom
                            }

                            for (int r = 0; r < numNewRows; r++) { //goes through all rows
                                Row newRow = newSheet.createRow(r); //creates rows
                                for (int c = 0; c < cols; c++) { //through all columns, same as original
                                    Cell newCell = newRow.createCell(c); //creates cells
                                    if (r < startingComCell.getRowIndex()) { //writes header cells, same as original
                                        Cell oldCell = sheet.getRow(r).getCell(c); //gets cell from original sheet
                                        if (oldCell != null) {
                                            //sets same value to new file
                                            oldCell.setCellType(Cell.CELL_TYPE_STRING);
                                            newCell.setCellValue(oldCell.getStringCellValue());
                                        } else {
                                            newCell.setCellValue(""); //if cell is empty, make empty cell
                                        }

                                    } else {
                                        Part p = parts.get(r - startingComCell.getRowIndex()); //goes through part list

                                        if (p != null) {
                                            if (c == startingDesCell.getColumnIndex()) {
                                                String newCellCons = "";
                                                for (int i = 0; i < p.getDesignators().size(); i++) {
                                                    newCellCons += p.getDesignators().get(i); //writes designators
                                                    if (i < p.getDesignators().size() - 1) {
                                                        newCellCons += ", "; //adds commas after designators as long as it is not the last element
                                                    }
                                                }

                                                newCell.setCellValue(newCellCons); //sets value of cell using constructed string

                                            } else if (c == startingComCell.getColumnIndex()) {
                                                newCell.setCellValue(p.getPartNumber()); //writes part number
                                            } else if (c == startingQtyCell.getColumnIndex()) {
                                                newCell.setCellValue(p.getDesignators().size()); //writes quantity by using size of designator list
                                            } else {
                                                if (p.getOtherInfo() != null) {
                                                    if (p.getOtherInfo().get(c) != null) {
                                                        newCell.setCellValue(p.getOtherInfo().get(c)); //writes other info
                                                    } else {
                                                        newCell.setCellValue(""); //leaves blank if there is nothing
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            try (FileOutputStream outputStream = new FileOutputStream(fileTwoDirField.getText() + File.separator +fieldTwo)) { //makes output stream for second file
                                newBook.write(outputStream); //writes the workbook
                                System.out.println(fileTwoDirField.getText() + File.separator + fieldTwo);

                                Desktop.getDesktop().open(new File(fileTwoDirField.getText() + File.separator + fieldTwo)); //opens excel file automatically by getting its absolute path
                            } catch (Exception err) {
                                err.printStackTrace();
                            }
                            /**All possible errors**/
                        } catch (Exception excep) {
                            excep.printStackTrace();
                        }
                    } catch (Exception exe) {
                        exe.printStackTrace();
                        JOptionPane.showMessageDialog(null, "Error");
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Wrong File Format");
                }
            } else {
                JOptionPane.showMessageDialog(null, "Invalid File");
            }
        }
    }

    /*getLetters
     *Method to get letters from cell adress
     * @param String, the cell address
     * @return String, the letters
     */
    public String getLetters(String cellAddress) {
        String letters = "";
        for (int i = 0; i < cellAddress.length(); i++) {
            if (Character.isLetter(cellAddress.charAt(i))) {
                letters += cellAddress.charAt(i);
            }
        }
        return letters;
    }

    /*getNumbers
     *Method to get numbers from cell adress
     * @param String, the cell address
     * @return String, the numbers
     */
    public String getNumbers(String cellAddress) {
        String numbers = "";
        for (int i = 0; i < cellAddress.length(); i++) {
            if (Character.isDigit(cellAddress.charAt(i))) {
                numbers += cellAddress.charAt(i);
            }
        }
        return numbers;
    }

    public class FileOneButListener implements ActionListener { //listener to open file chooser gui
        @Override
        public void actionPerformed(ActionEvent e) {
            FileChooserGUI fcg;
            if (directory != null) {
                fcg = new FileChooserGUI(directory); //makes gui starting from location of last file selected
            } else {
                fcg = new FileChooserGUI(); //makes gui starting from default location
            }
            fileOneString = fcg.getPath(); //gets string of chosen file's path
            fileOneField.setText(fileOneString); //sets the file name field to the path chosen
            directory = fcg.getDir(); //sets new directory for next time file chooser is opened
        }
    }

    public class FileTwoButtonListener implements ActionListener {

        @Override
        public void actionPerformed(ActionEvent e) {
            JFileChooser fc = new JFileChooser();
            if (directoryTwo != null){
                fc.setCurrentDirectory(directoryTwo);
            }
            fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            fc.showSaveDialog(null);
            fileTwoDirField.setText(fc.getSelectedFile().toString());
            directoryTwo = fc.getSelectedFile();
        }
    }

    public class HelpButtonListener implements ActionListener{

        /**
         * Invoked when an action occurs.
         *
         * @param e
         */
        @Override
        public void actionPerformed(ActionEvent e) {
            final String newline = "\n";
            area = new JTextArea(20, 20);
            area.setEditable(false);
            area.append("BOMOrganizer.jar" + newline);
            area.append("-------------------------------------------------------------------------------------------------------------------------------" + newline);
            area.append("Rearrange Customer Supplied Bill of Material.");
            area.append("Use the select button to enter an Excel file. (.xlsx and .xls both work)" + newline);
            area.append("Enter starting Component ID, Part Number, and Quantity cells. Ex Q7. Input can be lower or upper case."+newline);
            area.append("Use the select button to enter the new directory." + newline);
            area.append("Type in the name of the new file. If you do not include an extension, it will automatically be .xlsx." + newline);
            area.append("After clicking the Start button, the program will go through all part numbers and group all component IDs together in one line." + newline);
            area.append("It will also add all quantities together." + newline);
            area.append("The program will then save the Excel file and then open it automatically." + newline);
            panel.add(area);
            panel.revalidate();
            panel.repaint();
            helpButton.setEnabled(false);
        }
    }
}

