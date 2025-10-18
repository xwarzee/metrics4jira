package com.jira.explorer.ui;

import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.jira.explorer.model.JiraConfig;
import com.jira.explorer.model.JiraIssue;
import com.jira.explorer.service.JiraApiClient;
import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Orientation;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.scene.text.Font;
import javafx.scene.text.FontWeight;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;

/**
 * Main controller for the Jira JQL Explorer UI
 */
public class MainViewController {
    private static final Logger logger = LoggerFactory.getLogger(MainViewController.class);

    private final BorderPane root;
    private final TextField jqlTextField;
    private final TextArea resultTextArea;
    private final ListView<JiraIssue> issueListView;
    private final TableView<Map.Entry<String, Object>> fieldTableView;
    private final Label statusLabel;
    private final Button searchButton;
    private final Button configButton;

    private JiraApiClient jiraClient;
    private ObservableList<JiraIssue> issues;
    private JsonObject fieldMetadata;

    public MainViewController() {
        this.root = new BorderPane();
        this.issues = FXCollections.observableArrayList();

        // Initialize UI components
        this.jqlTextField = new TextField();
        this.resultTextArea = new TextArea();
        this.issueListView = new ListView<>(issues);
        this.fieldTableView = new TableView<>();
        this.statusLabel = new Label("Not connected to Jira");
        this.searchButton = new Button("Search");
        this.configButton = new Button("Configure");

        setupUI();
        setupEventHandlers();
    }

    private void setupUI() {
        root.setPadding(new Insets(10));

        // Top - Configuration and JQL Query
        VBox topSection = createTopSection();
        root.setTop(topSection);

        // Center - Split pane with issues list and field explorer
        SplitPane centerPane = createCenterSection();
        root.setCenter(centerPane);

        // Bottom - Status bar
        HBox bottomSection = createBottomSection();
        root.setBottom(bottomSection);
    }

    private VBox createTopSection() {
        VBox topSection = new VBox(10);
        topSection.setPadding(new Insets(0, 0, 10, 0));

        // Title
        Label titleLabel = new Label("Jira JQL Explorer - Interactive Whiteboard");
        titleLabel.setFont(Font.font("System", FontWeight.BOLD, 18));

        // Configuration bar
        HBox configBar = new HBox(10);
        configBar.getChildren().addAll(configButton, statusLabel);
        statusLabel.setStyle("-fx-text-fill: #666;");

        // JQL Query input
        HBox queryBar = new HBox(10);
        Label jqlLabel = new Label("JQL Query:");
        jqlLabel.setMinWidth(80);
        jqlTextField.setPromptText("Enter JQL query (e.g., project = MYPROJECT AND status = Open)");
        HBox.setHgrow(jqlTextField, Priority.ALWAYS);
        searchButton.setDefaultButton(true);
        searchButton.setStyle("-fx-background-color: #0052CC; -fx-text-fill: white;");
        queryBar.getChildren().addAll(jqlLabel, jqlTextField, searchButton);

        topSection.getChildren().addAll(titleLabel, configBar, queryBar);
        return topSection;
    }

    private SplitPane createCenterSection() {
        SplitPane splitPane = new SplitPane();
        splitPane.setOrientation(Orientation.HORIZONTAL);
        splitPane.setDividerPositions(0.3);

        // Left - Issues list
        VBox leftPane = new VBox(5);
        Label issuesLabel = new Label("Issues");
        issuesLabel.setFont(Font.font("System", FontWeight.BOLD, 14));
        issueListView.setPlaceholder(new Label("No issues loaded. Enter a JQL query and click Search."));
        VBox.setVgrow(issueListView, Priority.ALWAYS);
        leftPane.getChildren().addAll(issuesLabel, issueListView);
        leftPane.setPadding(new Insets(5));

        // Right - Split pane for field explorer and raw JSON
        SplitPane rightSplitPane = new SplitPane();
        rightSplitPane.setOrientation(Orientation.VERTICAL);
        rightSplitPane.setDividerPositions(0.6);

        // Top right - Field explorer
        VBox fieldExplorerPane = createFieldExplorerPane();

        // Bottom right - Raw JSON viewer
        VBox jsonViewerPane = createJsonViewerPane();

        rightSplitPane.getItems().addAll(fieldExplorerPane, jsonViewerPane);

        splitPane.getItems().addAll(leftPane, rightSplitPane);
        return splitPane;
    }

    private VBox createFieldExplorerPane() {
        VBox fieldPane = new VBox(5);
        Label fieldLabel = new Label("Field Explorer");
        fieldLabel.setFont(Font.font("System", FontWeight.BOLD, 14));

        // Setup field table
        TableColumn<Map.Entry<String, Object>, String> fieldNameCol = new TableColumn<>("Field Name");
        fieldNameCol.setCellValueFactory(param -> {
            String fieldId = param.getValue().getKey();
            String displayName = getFieldDisplayName(fieldId);
            return new SimpleStringProperty(displayName);
        });
        fieldNameCol.setPrefWidth(200);

        TableColumn<Map.Entry<String, Object>, String> fieldValueCol = new TableColumn<>("Value");
        fieldValueCol.setCellValueFactory(param -> {
            Object value = param.getValue().getValue();
            String displayValue = value != null ? value.toString() : "";
            // Truncate long values
            if (displayValue.length() > 200) {
                displayValue = displayValue.substring(0, 197) + "...";
            }
            return new SimpleStringProperty(displayValue);
        });
        fieldValueCol.setPrefWidth(400);

        fieldTableView.getColumns().addAll(fieldNameCol, fieldValueCol);
        fieldTableView.setPlaceholder(new Label("Select an issue to view its fields"));
        VBox.setVgrow(fieldTableView, Priority.ALWAYS);

        // Add copy button
        Button copyFieldButton = new Button("Copy Field Value");
        copyFieldButton.setOnAction(e -> copySelectedFieldValue());

        fieldPane.getChildren().addAll(fieldLabel, fieldTableView, copyFieldButton);
        fieldPane.setPadding(new Insets(5));
        return fieldPane;
    }

    private VBox createJsonViewerPane() {
        VBox jsonPane = new VBox(5);
        Label jsonLabel = new Label("Raw JSON");
        jsonLabel.setFont(Font.font("System", FontWeight.BOLD, 14));

        resultTextArea.setEditable(false);
        resultTextArea.setWrapText(true);
        resultTextArea.setFont(Font.font("Courier New", 12));
        VBox.setVgrow(resultTextArea, Priority.ALWAYS);

        Button copyJsonButton = new Button("Copy JSON");
        copyJsonButton.setOnAction(e -> copyJsonToClipboard());

        jsonPane.getChildren().addAll(jsonLabel, resultTextArea, copyJsonButton);
        jsonPane.setPadding(new Insets(5));
        return jsonPane;
    }

    private HBox createBottomSection() {
        HBox bottomSection = new HBox(10);
        bottomSection.setPadding(new Insets(10, 0, 0, 0));

        Label infoLabel = new Label("Total Issues: 0");
        infoLabel.setId("infoLabel");
        bottomSection.getChildren().add(infoLabel);

        return bottomSection;
    }

    private void setupEventHandlers() {
        searchButton.setOnAction(e -> executeSearch());
        configButton.setOnAction(e -> showConfigDialog());

        issueListView.getSelectionModel().selectedItemProperty().addListener(
            (observable, oldValue, newValue) -> {
                if (newValue != null) {
                    displayIssueDetails(newValue);
                }
            }
        );

        // Load field metadata in background
        loadFieldMetadata();
    }

    private void executeSearch() {
        String jql = jqlTextField.getText().trim();
        if (jql.isEmpty()) {
            showAlert("JQL Query Required", "Please enter a JQL query.");
            return;
        }

        if (jiraClient == null) {
            showAlert("Not Connected", "Please configure Jira connection first.");
            showConfigDialog();
            return;
        }

        searchButton.setDisable(true);
        statusLabel.setText("Searching...");

        Thread searchThread = new Thread(() -> {
            try {
                var results = jiraClient.searchIssues(jql, 0, jiraClient.getConfig().getMaxResults());
                Platform.runLater(() -> {
                    issues.clear();
                    issues.addAll(results);
                    updateInfoLabel();
                    statusLabel.setText("Search completed. Found " + results.size() + " issues.");
                    searchButton.setDisable(false);
                });
            } catch (Exception ex) {
                logger.error("Search failed", ex);
                Platform.runLater(() -> {
                    showAlert("Search Failed", "Error: " + ex.getMessage());
                    statusLabel.setText("Search failed");
                    searchButton.setDisable(false);
                });
            }
        });
        searchThread.setDaemon(true);
        searchThread.start();
    }

    private void displayIssueDetails(JiraIssue issue) {
        // Update field table
        ObservableList<Map.Entry<String, Object>> fieldEntries =
            FXCollections.observableArrayList(issue.getFlattenedFields().entrySet());
        fieldTableView.setItems(fieldEntries);

        // Update raw JSON
        JsonObject fields = issue.getFields();
        resultTextArea.setText(prettyPrintJson(fields));
    }

    private void showConfigDialog() {
        ConfigDialog dialog = new ConfigDialog(jiraClient != null ? jiraClient.getConfig() : new JiraConfig());
        dialog.showAndWait().ifPresent(config -> {
            jiraClient = new JiraApiClient(config);
            if (jiraClient.testConnection()) {
                statusLabel.setText("Connected to " + config.getJiraUrl());
                statusLabel.setStyle("-fx-text-fill: green;");
                loadFieldMetadata();
            } else {
                statusLabel.setText("Connection failed!");
                statusLabel.setStyle("-fx-text-fill: red;");
                showAlert("Connection Failed", "Could not connect to Jira. Please check your credentials.");
            }
        });
    }

    private void loadFieldMetadata() {
        if (jiraClient == null) return;

        Thread metadataThread = new Thread(() -> {
            try {
                fieldMetadata = jiraClient.getFieldMetadata();
                logger.info("Field metadata loaded");
            } catch (Exception ex) {
                logger.error("Failed to load field metadata", ex);
            }
        });
        metadataThread.setDaemon(true);
        metadataThread.start();
    }

    private String getFieldDisplayName(String fieldId) {
        if (fieldMetadata != null && fieldMetadata.has(fieldId)) {
            return fieldMetadata.get(fieldId).getAsString() + " (" + fieldId + ")";
        }
        return fieldId;
    }

    private void copySelectedFieldValue() {
        Map.Entry<String, Object> selectedEntry = fieldTableView.getSelectionModel().getSelectedItem();
        if (selectedEntry != null) {
            String value = selectedEntry.getValue() != null ? selectedEntry.getValue().toString() : "";
            javafx.scene.input.Clipboard clipboard = javafx.scene.input.Clipboard.getSystemClipboard();
            javafx.scene.input.ClipboardContent content = new javafx.scene.input.ClipboardContent();
            content.putString(value);
            clipboard.setContent(content);
            statusLabel.setText("Field value copied to clipboard");
        }
    }

    private void copyJsonToClipboard() {
        String json = resultTextArea.getText();
        if (!json.isEmpty()) {
            javafx.scene.input.Clipboard clipboard = javafx.scene.input.Clipboard.getSystemClipboard();
            javafx.scene.input.ClipboardContent content = new javafx.scene.input.ClipboardContent();
            content.putString(json);
            clipboard.setContent(content);
            statusLabel.setText("JSON copied to clipboard");
        }
    }

    private void updateInfoLabel() {
        Label infoLabel = (Label) root.getBottom().lookup("#infoLabel");
        if (infoLabel != null) {
            infoLabel.setText("Total Issues: " + issues.size());
        }
    }

    private String prettyPrintJson(JsonObject json) {
        com.google.gson.GsonBuilder builder = new com.google.gson.GsonBuilder();
        builder.setPrettyPrinting();
        return builder.create().toJson(json);
    }

    private void showAlert(String title, String message) {
        Alert alert = new Alert(Alert.AlertType.WARNING);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }

    public BorderPane getRoot() {
        return root;
    }
}
