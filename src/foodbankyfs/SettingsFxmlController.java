/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package foodbankyfs;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.net.URL;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.ResourceBundle;
import java.util.Scanner;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.event.EventType;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.geometry.Insets;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.effect.DropShadow;
import javafx.scene.image.Image;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.Background;
import javafx.scene.layout.BackgroundFill;
import javafx.scene.layout.CornerRadii;
import javafx.scene.layout.GridPane;
import javafx.scene.paint.Color;
import javafx.scene.paint.Paint;
import javafx.stage.*;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * FXML Controller class
 *
 * @author IT
 */
public class SettingsFxmlController implements Initializable {

    @FXML private GridPane SettingsGridPane;    
    @FXML private Label currentFile;
    @FXML private Button browseButton;
    static private boolean fileChange;
    
    /**
     * Initializes the controller class.
     */
    @Override
    public void initialize(URL url, ResourceBundle rb) {

        // Indicates file path has not been changed
        fileChange = false;
        
        // Show a shadow around button to indicate it has been clicked
        browseButton.addEventHandler(MouseEvent.MOUSE_PRESSED, new EventHandler<MouseEvent>() {
        
            @Override
            public void handle(MouseEvent event) {
                browseButton.setEffect(new DropShadow());
            }
        });
        
        // Remove shadow when button has been released
        browseButton.addEventHandler(MouseEvent.MOUSE_RELEASED, new EventHandler<MouseEvent>() {
        
            @Override
            public void handle(MouseEvent event) {
                browseButton.setEffect(null);
            }
        });
        
        Scanner reader = null;
        
        // Open and read settings file
        try {
            reader = new Scanner(new File("./settings.txt"));
        } catch (FileNotFoundException e) {
            System.err.println(e);
        }
        
        while (reader.hasNext()) {
            String line = reader.nextLine();
            
            // Read current spreadsheet file path and display
            if (line.contains("spreadsheet")) {
                String parseLine[] = line.split(",");
                currentFile.setText(parseLine[1]);
            }
        }
        
    }    
    
    /*
    / Calls file browser so user can select spreadsheet holding data
    */
    public void browseForFile(ActionEvent event) {
        
        // Create file chooser object
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Open Spreadsheet File");
        Stage stage = (Stage) SettingsGridPane.getScene().getWindow();
        
        // Show file chooser 
        File file = fileChooser.showOpenDialog(stage);
        
        try {
            if (file != null) {

                // Attempt to open settings file
                File settings = new File("./settings.txt");
                
                // Create new settings file if it does not exist
                if (!settings.exists()) {
                    settings.createNewFile();
                }
                
                // Write selected file path location to settings file
                List<String> data = Arrays.asList("spreadsheet, "+file.getAbsolutePath());
                Path filePath = settings.toPath();
                Files.write(filePath, data, Charset.forName("UTF-8"));
                
                // Update current path on settings menu
                currentFile.setText(file.getPath());
                
                // Indicates file path has been changed
                fileChange = true;
                
            }
        } catch (IOException e) {
            System.err.println(e);
        }

    }
    
    /*
    * Returns the file change boolean which indicates whether or not
    * file path was recently altered
    */
    public boolean getFileChangeStatus() {
        return fileChange;
    }
}
