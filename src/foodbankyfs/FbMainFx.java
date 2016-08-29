/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package foodbankyfs;


import java.awt.event.ActionListener;
import java.awt.Font;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.List;
import java.util.Optional;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.animation.RotateTransition;
import javafx.application.Application;
import javafx.application.Platform;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXMLLoader;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Group;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.image.ImageView;
import javafx.scene.layout.Background;
import javafx.scene.layout.BackgroundFill;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Pane;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.paint.Paint;
import javafx.scene.text.TextAlignment;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.Window;
import javafx.stage.WindowEvent;
import javafx.util.Duration;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author IT
 */
public class FbMainFx extends Application {
    
    private TextField msrIDField;
    private Label idScanResult;
    private TextField msrIDEntryField;
    private Button manualIDInputButton;
    private Button resetSpreadsheetMonth;
    private Button resetSpreadsheetWeek;
    private Button cancelManualInput;
    private File spreadsheet;
    private List<String[]> data;
    private boolean visitedAlreadyThisWeek;
    private XSSFSheet wrksheet;
    private XSSFWorkbook wrkbook;
    private boolean weeklyResetPerformed;
    private boolean monthlyResetPerformed;
    private SettingsFxmlController controller;
    
    @Override
    public void start(Stage primaryStage) {
        
        // Initialize spreadsheet file that holds user data.
        data = initSpreadsheetFile();
        checkResetDate();
        
        // If no spreadsheet is found notify user that one must be linked in the settings menu
        if (data == null) {
            Alert alert = createConfirmationAlertMsg("Missing file", 
                    "No spreadsheet linked.", 
                    "Please link the appropriate spreadsheet to the system in settings.");
            alert.showAndWait();
        }
                
        // Override default close operation.
        primaryStage.setOnCloseRequest(new EventHandler<WindowEvent>() {
            @Override public void handle(WindowEvent t) {
                closeConfirmation();
                t.consume();
            }
        });
        
        // Whenever window gains focus clear the MSR field
        primaryStage.focusedProperty().addListener(new ChangeListener<Boolean>()
        {
            @Override
            public void changed(ObservableValue<? extends Boolean> ov, Boolean t, Boolean t1)
            {
                clearMsrField();
            }
        });
        
        // Create grid.
        GridPane grid = new GridPane();
        grid.setAlignment(Pos.CENTER);
        grid.setHgap(10);
        grid.setVgap(10);
        grid.setPadding(new Insets(25, 25, 25, 25));
        grid.setId("grid");
        
        // Create text field for reading MSR input
        msrIDField = new TextField();
        msrIDField.setAlignment(Pos.CENTER);    // Does not seem to center the text
        msrIDField.setMaxSize(260, 1);
        msrIDField.setMinSize(260, 1);
        
        // Retrieve icon image and create animation icon
        final ImageView spinner =
        new ImageView(FbConstants.ICON_URL);
        
        spinner.setFitHeight(20);
        spinner.setFitWidth(20);
        spinner.setVisible(false);
        grid.add(spinner, 2, 11);
        
        // setup some transition that rotates an icon for X seconds
        final RotateTransition rotateTransition = new RotateTransition(Duration.seconds(FbConstants.DELAY_FOR_CARD_SCAN), spinner);
        rotateTransition.setByAngle(90);
        
        // delay rotation 
        rotateTransition.setDelay(Duration.seconds(FbConstants.DELAY_FOR_CARD_SCAN));
    
        // Makes spinning icon visible and executes animation
        msrIDField.textProperty().addListener(new ChangeListener<String>() {
            @Override
            public void changed(ObservableValue<? extends String> ov, String t, String currentString) {
                spinner.setVisible(true);
                rotateTransition.playFromStart();
            }
        });
        
        // When animation is complete execute change handler
        rotateTransition.setOnFinished((finishHim) -> {
            // Detects when text is entered into textbox and performs function when it matches regex unless spreadsheet not linked.
            if (data != null) {    
                outputMsrScanResult(awaitMsrSwipe(null));
            }
            
            spinner.setVisible(false);
        });
        
        grid.setAlignment(Pos.CENTER);
        grid.add(msrIDField, 0, 11, 2, 1);
        
        // Create label for outputting registered or unregistered.
        idScanResult = new Label("");
        idScanResult.setTextAlignment(TextAlignment.CENTER);
        grid.add(idScanResult, 0, 13, 3, 1);
        
        // Initially invisible textfield for manual input of id
        msrIDEntryField = new TextField();
        msrIDEntryField.setMaxSize(120, 30);
        msrIDEntryField.setMinSize(120, 30);
        msrIDEntryField.setVisible(false);
        grid.add(msrIDEntryField, 0, 15);
        
        // Button to bring up another textfield for manual ID entry
        manualIDInputButton = new Button("Manual ID Input");
        manualIDInputButton.setOnAction(new EventHandler<ActionEvent>() 
        {
            @Override public void handle(ActionEvent e) 
            {
                // If manual id entry field is visible then submit text entry
                if (msrIDEntryField.isVisible()) {
                    String id = "";
                    
                    // Guard against blank entries
                    if (msrIDEntryField.getText() != null && !msrIDEntryField.getText().isEmpty()) 
                    {
                        if (spreadsheet != null) {
                            id = msrIDEntryField.getText().toLowerCase().trim();
                            outputMsrScanResult(awaitMsrSwipe(id));
                            msrIDEntryField.clear();
                            msrIDEntryField.setVisible(false);
                            cancelManualInput.setVisible(false);
                            manualIDInputButton.setText("Manual ID Input");
                            msrIDField.requestFocus();
                        } else {
                            Alert alert = createConfirmationAlertMsg("No spreadsheet linked.",
                                "No spreadsheet selected in settings", 
                                "A spreadsheet must be linked to the application using the settings menu.");
                            alert.showAndWait();
                        }
                    }
                    else {
                        Alert alert = createConfirmationAlertMsg("Invalid entry.",
                                "Manual ID entry field empty", 
                                "Manual ID entry field cannot be empty. Please enter in a valid ID value.");
                        alert.showAndWait();
                    }
                                        
                }
                // If id field is not visible make it visible and change text on button
                else {
                    msrIDEntryField.setVisible(true);
                    cancelManualInput.setVisible(true);
                    manualIDInputButton.setText("Submit ID");
                }
            }
        });
        manualIDInputButton.setAlignment(Pos.CENTER_RIGHT);
        grid.add(manualIDInputButton, 1, 15);

        // Create cancel button for manual ID input
        cancelManualInput = new Button("X");
        cancelManualInput.setBackground(new Background(new BackgroundFill(Paint.valueOf("Red"), null, null)));
        cancelManualInput.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            // Cancel manual entry and hide manual entry text field and reset entry button 
            public void handle(ActionEvent e) {
                msrIDEntryField.clear();
                msrIDEntryField.setVisible(false);
                cancelManualInput.setVisible(false);
                manualIDInputButton.setText("Manual ID Input");
                msrIDField.requestFocus();
                cancelManualInput.setVisible(false);
            }
        });
        cancelManualInput.setAlignment(Pos.BASELINE_LEFT);
        cancelManualInput.setVisible(false);
        grid.add(cancelManualInput, 2, 15);
        
        // Create and add reset buttons
        resetSpreadsheetMonth = new Button("Reset Monthly Totals");
        resetSpreadsheetMonth.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent e) {
                // Prompts user for confirmation before reset is performed
                Alert alert = createConfirmationAlertMsg("Monthly Reset",
                        "Attempting Monthly Reset.",
                        "The system automatically resets the monthly visit totals on the 1st of every month." +
                                "Resetting it using this button is required only if the system was not run on the 1st");
                Optional<ButtonType> result = alert.showAndWait();
                
                if (result.get() == ButtonType.OK) {
                    resetSpreadsheetTotals(true);
                }
                      
            }
        });
        grid.add(resetSpreadsheetMonth, 1, 16, 2, 1);
        
        resetSpreadsheetWeek = new Button("Reset Weekly Visits");
        resetSpreadsheetWeek.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent e) {
                Alert alert = createConfirmationAlertMsg("Monthly Reset",
                        "Attempting Weekly Reset.",
                        "The system automatically resets the weekly has visited value on Monday." +
                                "Resetting it using this button is required only if the system was not run on a Monday");
                Optional<ButtonType> result = alert.showAndWait();
                
                if (result.get() == ButtonType.OK) {
                    resetSpreadsheetTotals(false);
                }
 
            }
        });
        grid.add(resetSpreadsheetWeek, 1, 17, 2, 1);
        
        // Create scene and start application.
        VBox root = new VBox();
        Scene scene = new Scene(root, 325, 375);   
        scene.getStylesheets().add(FbMainFx.class.getResource("Style.css").toExternalForm());
        
        // Initialize menu bar and add it to stage
        initMenu(scene);
        
        // Add gridpane to stage
        root.getChildren().add(grid);
        
        primaryStage.setTitle(FbConstants.APP_TITLE);
        primaryStage.setScene(scene);
        primaryStage.setResizable(false);
        primaryStage.show();
        
        
        
    }
    
    /*
     * Check settings file and load spreadsheet file path if it exists.
    */
    private List<String[]> initSpreadsheetFile() {
        
        // Initialize application settings file.
        File settings = new File("./settings.txt");
        
        try {
            // If settings file does not exist create it.
            if (settings.exists()) {
                Scanner reader = new Scanner(settings);
                
                // Read through each line of settings file to find file path.
                while (reader.hasNext()) {
                    String line = reader.nextLine();
                    
                    if (line.contains("spreadsheet")) {
                        String settingsParse[] = line.split(",");
                        
                        // Set spreadsheet file to file in path
                        spreadsheet = new File(settingsParse[1].trim());
                                                
                        // If file is empty then reset the value to null
                        if (spreadsheet.length() == 0) {
                            spreadsheet = null;
                        }

                    }
                    // Store weekly reset value
                    else if (line.contains("weeklyReset")) {
                        String settingsParse[] = line.split(",");
                        
                        weeklyResetPerformed = settingsParse[1].trim().equals("true");
                    }
                    // Store monthly reset value
                    else if (line.contains("monthlyReset")) {
                        String settingsParse[] = line.split(",");
                        
                        monthlyResetPerformed = settingsParse[1].trim().equals("true");
                    }
                }
                
                reader.close();
                
                if (spreadsheet != null) {
                    // Calls method to save contents of file to memory
                    return saveSpreadsheetData();
                }
                
            }
            else {
                settings.createNewFile();
                return null;
            }
        }
        catch (IOException e) {
            System.err.println(e);
        }
        
        return null;
    }
    
    /*
    * Saves contents of spreadsheet data into memory
    */
    private List<String[]> saveSpreadsheetData() {
        
        // Copy spreadsheet contents into memory
        try {

            // Initialize xls reading objects
            FileInputStream fileInputStream = new FileInputStream(spreadsheet);
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheet(FbConstants.SHEET_NAME);
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            List<String[]> tmpData = new ArrayList();

            // Save XSSF objects for rewrite
            wrksheet = worksheet;
            wrkbook = workbook;
            
            // Iterate through all rows in the sheet
            for (int rowNum = FbConstants.DATA_ROW_START_INDEX; rowNum < worksheet.getLastRowNum(); rowNum++) {
                
                // Initialize array that will store cell contents
                String values[] = new String[FbConstants.NUMBER_OF_COLUMNS];
                XSSFRow row = worksheet.getRow(rowNum);
                
                // Iterate through cells in each row and store values to an array
                for (int cellNum = 0; cellNum < FbConstants.NUMBER_OF_COLUMNS; cellNum++) {
                    XSSFCell cell = row.getCell(cellNum, Row.CREATE_NULL_AS_BLANK);
                      
                    String value = "";
                    
                    if (cell != null) {
                        
                        if (cell.getCellType() == XSSFCell.CELL_TYPE_FORMULA) {
                            evaluator.evaluateInCell(cell);
                        }
                        // If cell type is numeric convert the number value to a string
                        if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                            double tmpVal = cell.getNumericCellValue();
                            value = String.format("%.0f", tmpVal);
                        }
                        if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
                            value = cell.getStringCellValue().toLowerCase();
                            
                        }
                       
                    }
                    
                    // If a cell row has an empty ID do not include it in data
                    if (cellNum == 0 && value.equals("")) {
                        break;
                    }

                    // Initialize value to 0 if cell is empty
                    if (value.isEmpty()) {
                        
                        // If value is from email or notes field then put empty instead
                        if (cellNum == FbConstants.EMAIL_FIELD || cellNum == FbConstants.NOTES_FIELD) {
                            value = "empty";
                        }
                        else {
                            value = "0";
                        }

                    }
                    
                    // Store value in array
                    values[cellNum] = value;

                    
                }

                // Store array of values in list
                tmpData.add(values);
                
            }
            
            return tmpData;
            
        }catch (IOException e) {
            System.err.println(e);
        }
        
        return null;
    }
    
    /*
     * Initialize the menu bar and it's items.
    */
    private void initMenu(Scene scene) {
        MenuBar menuBar = new MenuBar();
        
        Menu menuMain = new Menu("Menu");
        
        // Clicking settings opens up a settings window.
        MenuItem menuSettings = new MenuItem("Settings");
        menuSettings.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent t) {
                openSettingsMenu(t);
            }
        });
                
        // Clicking exit generates a confirmation window for exiting the application.
        MenuItem menuExit = new MenuItem("Exit");
        menuExit.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent t) {
                closeConfirmation();
                t.consume();
            }
        });
        
        Menu menuHelp = new Menu("Help");
        
        // Clicking about displays application specific creator data
        MenuItem menuAbout = new MenuItem("About");
        menuAbout.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent t) {
                Alert alert = createConfirmationAlertMsg("About", 
                        "Author: Dan To\n"+
                        "Date: March 16, 2016", 
                        "Software reads and writes to local spreadsheet file "
                                + "and uses data in file to compare against text or MSR entry");
                alert.showAndWait();
            }
        });
        
        // Clicking info displays an application use manual
        MenuItem menuInfo = new MenuItem("Info");
        menuInfo.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent t) {
                Alert alert = createConfirmationAlertMsg("Information Manual",
                        "Troubleshooting Tips", 
                        "-If this is your first time using the program remember a spreadsheet file needs to be linked to the program in settings before it will work\n" +
                        "-Make sure the application window is in focus when attempting to swipe MSR cards\n" +
                        "-If the swipe result is an unexpected value, attempt the swipe again or minimize the window and open it again" );
                alert.showAndWait();
            }
        });
        
        // Add menu options to menu bar and set properties
        menuMain.getItems().addAll(menuSettings, menuExit);
        menuHelp.getItems().addAll(menuInfo, menuAbout);
        
        menuBar.prefWidthProperty().bind(scene.widthProperty());
        menuBar.getMenus().addAll(menuMain, menuHelp);
        
        ((VBox) scene.getRoot()).getChildren().add(menuBar);
    }
    
    /*
     * Method generates new scene for the settings menu.
    */
    private void openSettingsMenu(ActionEvent event) {
        
        FXMLLoader root;
        try {
            
            // Specifies FXML to load
            root = new FXMLLoader(getClass().getResource("Settings.fxml"), null);
            
            // Create stage for settings menu
            Stage settingsStage = new Stage();
            settingsStage.setTitle("Settings");
            settingsStage.setScene(new Scene((Pane) root.load()));
            
            // Override default close operation.
            settingsStage.setOnCloseRequest(new EventHandler<WindowEvent>() {
                @Override public void handle(WindowEvent t) {
                    
                    // When settings menu is closed if file change performed update data variable
                    if (controller.getFileChangeStatus()) {
                        data = initSpreadsheetFile();
                    }

                    settingsStage.close();
                    t.consume();
                }
            });
            
            // Initialize controller class for Settings.fxml
            controller = root.<SettingsFxmlController>getController();
            
            // Display settings scene
            settingsStage.show();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    /*
     * Parses York University student cards MSR track data and returns student_id.
     * As of the time writing this software track data format is:
     * %[student_id]?;5152900701[7digit_id_of_card]? LENGTH = 30
     */
    private String parseMsrID(String data) {
        //Splits string into student_id and card_id
        String fields[] = data.split(";");
        
        //Removes % and ? from student_id
        String studentID = fields[0].substring(1, fields[0].length() - 1);
        
        return studentID;
    }
    
    /*
     * Waits for user to swipe MSR card.
    */
    public int awaitMsrSwipe(String manualID) {
        
        // Get text from text field
        String IDField = msrIDField.getText();
        
        // Ensures text field is not null
        if (IDField != null || manualID != null) {

            // Ensures at least 29 characters have been entered (length of MS content)
            if (IDField.length() > 29 || manualID != null) 
            {
                // If text in field matches this pattern characters;characters
                if (IDField.matches(".+;.+\\?$") || manualID != null) {
                    
                    // Parse entry to get ID
                    String id = "";
                    
                    // Assign id based on MSR swipe or manual entry
                    if (manualID == null) {
                        id = parseMsrID(IDField.trim());
                    }
                    else {
                        id = manualID;
                    }
                    
                    // Checks if ID is in data
                    boolean result = searchIdInFile(id);
                    
                    // Method returns true or false depending on id search result
                    if (result) {
                        return FbConstants.ID_FOUND;
                    }
                    else {
                        return FbConstants.ID_NOT_FOUND;
                    }
                    
                }
                else {
                    
                    // If text does not match expected sequence clear field
                    clearMsrField();

                }
            }
            
        }
        
        return FbConstants.ID_NOT_VALID;
    }
    
    /*
    * Takes id parameter and searches memory contents to see if it is present
    * returns TRUE if id is present in memory
    */
    private boolean searchIdInFile(String id) {
        
        visitedAlreadyThisWeek = false;
        
        for (int rowNum = 0; rowNum < data.size(); rowNum++) {
            String row[] = data.get(rowNum);
            String rowID = "";
            
            // Guards against missing or incorrect student ids in the spreadsheet
            if (row[FbConstants.STUDENT_ID_FIELD] == null)
            {
                rowID = "0";
            }
            else {
                rowID = row[FbConstants.STUDENT_ID_FIELD];
            }

            if (rowID.equals(id)) {
                
                // Checks to make sure user has not visited this week
                if (row[FbConstants.VISITED_THIS_WEEK_FIELD].equals("0")) {
                    
                    // Changes user visited this week flag to 1
                    row[FbConstants.VISITED_THIS_WEEK_FIELD] = "1";
                    
                    // Increments number of time user has visited this month
                    if (row[FbConstants.MONTLY_VISIT_TOTAL_FIELD] == null) {
                        row[FbConstants.MONTLY_VISIT_TOTAL_FIELD] = "1";
                    }
                    else {
                        row[FbConstants.MONTLY_VISIT_TOTAL_FIELD] = 
                                String.format("%d", Integer.parseInt(row[FbConstants.MONTLY_VISIT_TOTAL_FIELD]) + 1);
                    }
                    
                    // Update row data in data list
                    data.remove(rowNum);
                    data.add(row);
                    
                    // Update XSSF object with new sheet data
                    XSSFRow sheetRow = wrksheet.getRow(rowNum + FbConstants.DATA_ROW_START_INDEX);
                    sheetRow.getCell(FbConstants.VISITED_THIS_WEEK_FIELD, Row.CREATE_NULL_AS_BLANK).setCellValue(row[FbConstants.VISITED_THIS_WEEK_FIELD]);
                    sheetRow.getCell(FbConstants.MONTLY_VISIT_TOTAL_FIELD, Row.CREATE_NULL_AS_BLANK).setCellValue(row[FbConstants.MONTLY_VISIT_TOTAL_FIELD]);
                    
                    // Write new sheet data to local spreadsheet
                    FileOutputStream out;
                    try 
                    {
                        out = new FileOutputStream(spreadsheet);
                        wrkbook.write(out);
                        out.close();
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(FbMainFx.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        Logger.getLogger(FbMainFx.class.getName()).log(Level.SEVERE, null, ex);
                    }
                                        
                    return true;
                }
                else {
                    
                    // Set flag for user if they have already visited this week
                    visitedAlreadyThisWeek = true;
                    return false;
                }
                
            }
        }
        
        return false;
    }
    
    /*
     * Outputs MSR card ID scan result to idScanResult label.
    */
    public void outputMsrScanResult(int swipeCheck) {
                
        boolean swipeResult = false;
        boolean invalidMsr = false;
        
        // Determines output result based on result of swipe check of MSR field data
        switch (swipeCheck) {
            case FbConstants.ID_FOUND:
                swipeResult = true;
                break;
            case FbConstants.ID_NOT_FOUND:
                swipeResult = false;
                break;
            default:
                invalidMsr = true;
                break;
        }
        
        // If an invalid entry was accidently placed in the MSR field then don't output a result
        if (!invalidMsr) {
            if (swipeResult && !visitedAlreadyThisWeek) {           
                // Output user result as Registered
                idScanResult.setTextFill(Color.CHARTREUSE);
                idScanResult.setAlignment(Pos.CENTER);
                idScanResult.setStyle("-fx-font: 36 arial;");
                idScanResult.setText("   -Registered-");
            }
            else if (!swipeResult && visitedAlreadyThisWeek) {
                // Output user result as Weekly Limit Reached
                idScanResult.setTextFill(Color.DARKORANGE);
                idScanResult.setStyle("-fx-font: 26 arial;");
                idScanResult.setText("Weekly Limit Reached");
            } else {
                // Output user result as Unregistered
                idScanResult.setTextFill(Color.CRIMSON);
                idScanResult.setStyle("-fx-font: 32 arial;");
                idScanResult.setText("  Not Registered");
            }
        }
        
        clearMsrField();

    }
    
    /*
    / Checks date for time and performs spreadsheet total reset when date is a Monday
    / or day of the month is the 1st
    */
    private void checkResetDate() {
        
        // Grab system's current date
        Calendar calendar = Calendar.getInstance();

        SimpleDateFormat format = new SimpleDateFormat("yyyy/MM/dd");
        String date = format.format(calendar.getTime());
        boolean isMonday = calendar.get(Calendar.DAY_OF_WEEK) == Calendar.MONDAY;
        String day = date.split("/")[2].trim();
        
        // Perform a weekly reset when day of the week is a Monday
        if (isMonday && !weeklyResetPerformed) {
            // Parameter indicates whether the reset is a monthly reset
            resetSpreadsheetTotals(false);
        }
        
        // Perform a monthly reset when day of the month is 1
        if (day.equals("01") && !monthlyResetPerformed) {
            resetSpreadsheetTotals(true);
        }
        
        // If it's niether the first or a Monday then reset the flag values
        if (!day.equals("01") || !isMonday) {
            clearResetValues(day.equals("01"), isMonday);
        }
    }
    
    /*
    / Resets all the values in either weekly visit or monthly visit total based on passed boolean value
    / TRUE for monthly reset, FALSE for weekly reset
    */
    private void resetSpreadsheetTotals(boolean monthlyReset) {
        
        if (spreadsheet != null) {
            
            FileOutputStream fileOutputStream;
            
            try 
            {
            
                fileOutputStream = new FileOutputStream(spreadsheet);
                
                // Iterate through spreadsheet object and reset values
                for (int i = FbConstants.DATA_ROW_START_INDEX; i < wrksheet.getLastRowNum(); i++) 
                {
                    Row row = wrksheet.getRow(i);

                    // If monthly reset set all monthly visit totals to 0, else set weekly visit value to 0
                    if (monthlyReset) {
                        row.getCell(FbConstants.MONTLY_VISIT_TOTAL_FIELD, Row.CREATE_NULL_AS_BLANK).setCellValue("0");
                    }
                    else {
                        row.getCell(FbConstants.VISITED_THIS_WEEK_FIELD, Row.CREATE_NULL_AS_BLANK).setCellValue("0");
                    }
                }
            
                // Write changes to spreadsheet file
                wrkbook.write(fileOutputStream);
                fileOutputStream.close();
                
                // Write reset performed status to settings file
                if (monthlyReset) {
                    // Store data to write as an iterable list
                    List<String> datum = Arrays.asList("spreadsheet, "+spreadsheet.getAbsolutePath()+System.lineSeparator()
                        + "monthlyReset, true"+System.lineSeparator()
                        + "weeklyReset, "+weeklyResetPerformed);
                    Files.write((new File("./settings.txt")).toPath(), datum, Charset.forName("UTF-8"));
                }
                else {
                    List<String> datum = Arrays.asList("spreadsheet, "+spreadsheet.getAbsolutePath()+System.lineSeparator()
                        + "monthlyReset, "+monthlyResetPerformed+System.lineSeparator()
                        + "weeklyReset, true");
                    Files.write((new File("./settings.txt")).toPath(), datum, Charset.forName("UTF-8"));
                }
                
                // Refresh spreadsheet data
                data = initSpreadsheetFile();
                
            } catch (FileNotFoundException ex) {
                Logger.getLogger(FbMainFx.class.getName()).log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                Logger.getLogger(FbMainFx.class.getName()).log(Level.SEVERE, null, ex);
            }

        }
        
        // Return focus to text box
        if (msrIDField != null) {
            msrIDField.requestFocus();
        }
    }
    
    /*
     * Method clears the montly and weekly reset guards when it is not monday or not the beginning of the month
    */
    private void clearResetValues(boolean newMonth, boolean newWeek) {
        List<String> datum;
        
        if (spreadsheet != null) 
        {
            // If it's not a monday and a weekly rest has been performed then reset the value to false
            if (!newWeek && weeklyResetPerformed) {
                weeklyResetPerformed = false;
            }
            // If it's not the 1st of a month and the monthly reset has been performed then reset the value to false
            else if (!newMonth && monthlyResetPerformed) {
                monthlyResetPerformed = false;
            }

            datum = Arrays.asList("spreadsheet, "+spreadsheet.getAbsolutePath()+System.lineSeparator()
                            + "monthlyReset, "+monthlyResetPerformed+System.lineSeparator()
                            + "weeklyReset, "+weeklyResetPerformed);

            try {
                Files.write((new File("./settings.txt")).toPath(), datum, Charset.forName("UTF-8"));
            }
            catch (IOException e) {
                System.err.println(e);
            }
        }
    }
    
    /*
     * Creates an alert dialog box with specified parameters.
    */
    private Alert createConfirmationAlertMsg(String title, String header, String body) {
        Alert alert = new Alert(AlertType.CONFIRMATION);
        alert.setTitle(title);
        alert.setHeaderText(header);
        alert.setContentText(body);
        
        return alert;
    }
    
    // Method clears the MSR text field
    private void clearMsrField() {
        this.msrIDField.clear();
        this.msrIDField.setText("");
    }
    
    /*
     * Creates a dialog that requests user to confirm closing the application.
    */
    private void closeConfirmation() {
        
        Alert alert = createConfirmationAlertMsg("Confirmation", 
                        "You are attempting to close the application.",
                        "Are you sure you wish to exit?");
        Optional<ButtonType> result = alert.showAndWait();
                
        if (result.get() == ButtonType.OK){
            // ... user chose OK
            Platform.exit();
        }
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        launch(args);
        
    }
    
}
