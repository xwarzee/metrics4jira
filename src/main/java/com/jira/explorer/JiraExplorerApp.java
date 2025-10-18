package com.jira.explorer;

import com.jira.explorer.ui.MainViewController;
import javafx.application.Application;
import javafx.scene.Scene;
import javafx.stage.Stage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Main application class for Jira JQL Explorer
 * Interactive whiteboard for exploring Jira issues and their fields
 */
public class JiraExplorerApp extends Application {
    private static final Logger logger = LoggerFactory.getLogger(JiraExplorerApp.class);

    @Override
    public void start(Stage primaryStage) {
        try {
            logger.info("Starting Jira JQL Explorer application");

            MainViewController controller = new MainViewController();
            Scene scene = new Scene(controller.getRoot(), 1200, 800);

            // Add CSS styling
            scene.getStylesheets().add(getClass().getResource("/styles.css").toExternalForm());

            primaryStage.setTitle("Jira JQL Explorer - Interactive Whiteboard");
            primaryStage.setScene(scene);
            primaryStage.setMinWidth(800);
            primaryStage.setMinHeight(600);
            primaryStage.show();

            logger.info("Application started successfully");
        } catch (Exception e) {
            logger.error("Failed to start application", e);
            throw new RuntimeException(e);
        }
    }

    @Override
    public void stop() {
        logger.info("Stopping Jira JQL Explorer application");
    }

    public static void main(String[] args) {
        logger.info("Launching Jira JQL Explorer");
        launch(args);
    }
}
